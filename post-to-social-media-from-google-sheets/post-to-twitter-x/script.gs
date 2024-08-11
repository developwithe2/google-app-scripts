// GLOBAL VARIABLES
const scriptProps = PropertiesService.getScriptProperties();
const workbook = SpreadsheetApp.getActiveSpreadsheet();
const timesOfDay = {
  "Midnight": [0,3],
  "Early Morning": [3,6],
  "Morning": [6,9],
  "Late Morning": [9,12],
  "Early Afternoon": [12,15],
  "Afternoon": [15,18],
  "Evening": [18,21],
  "Late Evening": [21,24]
};

// POST TO TWITTERX
function postToTwitterX(scriptProps, workbook, timesOfDay) {
  const tx = {
    appName: 'TwitterX',
    clientId: scriptProps.getProperty('txClientId'),
    clientSecret: scriptProps.getProperty('txClientSecret'),
    accessToken: scriptProps.getProperty('txAccessToken'),
    refreshToken: scriptProps.getProperty('txRefreshToken'),
    expiration: parseInt(scriptProps.getProperty('txExpiresOn'), 10),
    baseUrl: scriptProps.getProperty('txBaseUrl'),
    endpointToken: scriptProps.getProperty('txEndpointToken'),
    endpointTweets: scriptProps.getProperty('txEndpointTweets'),
  };

  // CONFIRM ACTIVE TOKENS or REFRESH TOKENS
  const twitterTokens = initializeTwitterXTokens(scriptProps, tx);
  if (!twitterTokens) {
    Logger.log('ERROR: Did not receive tokens. Aborting execution.');
    return;
  };

  try {
    const columns = {
      status: columnIndex(tx.appName, 'status'),
      post_date: columnIndex(tx.appName, 'post_date'),
      time_of_day: columnIndex(tx.appName, 'time_of_day'),
      media_type: columnIndex(tx.appName, 'media_type'),
      caption: columnIndex(tx.appName, 'caption'),
      poll_options: columnIndex(tx.appName, 'poll_options'),
      poll_duration: columnIndex(tx.appName, 'poll_duration'),
      post_id: columnIndex(tx.appName, 'post_id'),
    };

    // GET ARRAY OF ALL POSTS WITHIN RANGE/TABLE
    const namedRange = workbook.getRangeByName(tx.appName);
    const postRows = namedRange.getValues();

    // SET FILTER PARAMETERS BASED ON DATE & TIME OF DAY
    const now = new Date();
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const currentHour = now.getHours();
    const currentTimeOfDay = Object.keys(timesOfDay).find(period => {
      const [start, end] = timesOfDay[period];
      return currentHour >= start && currentHour < end;
    });

    // FILTER POSTS BY PARAMETERS
    const filteredPosts = postRows.filter(row => {
      const postDate = Utilities.formatDate(new Date(row[columns.post_date]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const postTimeOfDay = row[columns.time_of_day];
      return postDate === today && postTimeOfDay === currentTimeOfDay;
    });

    if (filteredPosts.length === 0) {
      Logger.log('No TwitterX posts scheduled for today at this time.');
      return; // Terminate run
    };

    // POST & RECORD DATA FOR EACH POST WITHIN FILTERED ARRAY
    filteredPosts.forEach((row, index) => {
      const caption = row[columns.caption];
      const mediaType = row[columns.media_type];
      const pollOptions = row[columns.poll_options].split(',').map(option => option.trim());
      const pollDuration = row[columns.poll_duration];

      const postId = createTwitterXPost(tx, caption, mediaType, pollOptions, pollDuration, twitterTokens.access_token);
      Logger.log(postId);

      // RECORD POST DATA TO SPREADSHEET
      const rowIndex = namedRange.getRow() + index;
      workbook.getSheetByName(tx.appName).getRange(rowIndex, columns.post_id + 1).setValue(postId);
      workbook.getSheetByName(tx.appName).getRange(rowIndex, columns.status + 1).setValue('POSTED');
    });
  } catch (error) {
    Logger.log('An error occurred while running postToTwitterX:', error.message);
  }
}

// POST TWEET TO TWITTERX
function createTwitterXPost(tx, caption, mediaType, pollOptions, pollDuration, access_token) {
  try {
    // INITIATE PAYLOAD DATA
    let payload = {
      'text': caption
    };

    // ADD POST TYPE PAYLOAD DATA
    if (mediaType === 'POLL') {
      // POLL: options, duration_minutes
      payload = {
        ...payload,
        'poll': {
          'options': pollOptions,
          'duration_minutes': pollDuration
        }
      };
    }

    // SET OPTIONS DATA
    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: {
        Authorization: `Bearer ${access_token}`
      },
      muteHttpExceptions: true
    };

    // SET ENDPOINT
    const urlTweets = `${tx.baseUrl}${tx.endpointTweets}`;

    // SEND & RECORD POST SUBMISSION
    const createTwitterXPostRequest = UrlFetchApp.fetch(urlTweets, options);
    Logger.log(createTwitterXPostRequest);
    const createTwitterXPostResponse = JSON.parse(createTwitterXPostRequest.getContentText());

    // ERROR HANDLING
    if (createTwitterXPostRequest.getResponseCode() === 200 || createTwitterXPostRequest.getResponseCode() === 201) {
      Logger.log('Successfully posted Tweet.');
      return createTwitterXPostResponse.data.id;
    } else {
      Logger.log('ERROR: Unable to post tweet. Response code: ' + createTwitterXPostRequest.getResponseCode() + '. Response body: ' + createTwitterXPostRequest.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log('An error occurred while running createTwitterXPost:', error.message);
  }
}

// INITIALIZES OR UPDATES TOKENS
function initializeTwitterXTokens(scriptProps, tx) {
  let access_token = tx.accessToken;
  let refresh_token = tx.refreshToken;
  let expires_on = new Date(tx.expiration);

  // CONFIRM IF TODAY IS LATER THAN TOKEN EXPIRATION DATE
  if (new Date() > expires_on) {

    // REFRESH TOKENS (SEE ASSOCIATED FUNCTION)
    const response = refreshTwitterXTokens(tx, refresh_token);

    // UPDATE SCRIPT PROPERTIES WITH REFRESHED TOKENS
    if (response) {
      access_token = response.access_token;
      refresh_token = response.refresh_token;
      const now = new Date(); // Get datetime of token refresh
      expires_on = now.getTime() + (response.expires_in * 1000); // Set new token expiration date
      const expires_on_milliseconds = expires_on.toFixed(0); // Convert new token expiration date into milliseconds for storage
      scriptProps.setProperty('txAccessToken', access_token);
      scriptProps.setProperty('txRefreshToken', refresh_token);
      scriptProps.setProperty('txExpiresOn', expires_on_milliseconds);
    } else {
      return null;
    }
  } else {
    Logger.log('All tokens are currently active.');
  }
  // RETURN EXISTING TOKENS OR REFRESHED TOKENS
  return { access_token, refresh_token };
}

// REFRESH TOKENS
function refreshTwitterXTokens(tx, refresh_token) {

  const tokenPayload = {
    'refresh_token': refresh_token,
    'grant_type': "refresh_token",
    'client_id': tx.clientId
  };
  const basicAuthorizationHeader = Utilities.base64Encode(tx.clientId + ':' + tx.clientSecret);
  const tokenOptions = {
    method: 'POST',
    contentType: 'application/x-www-form-urlencoded',
    payload: Object.entries(tokenPayload).map(([key, value]) => encodeURIComponent(key) + '=' + encodeURIComponent(value)).join('&'),
    headers: {
      Authorization: 'Basic ' + basicAuthorizationHeader
    },
    muteHttpExceptions: true
  };

  // SET ENDPOINT
  const urlToken = `${tx.baseUrl}${tx.endpointToken}`;

  // SEND TOKEN REFRESH REQUEST
  const refreshTwitterXTokensRequest = UrlFetchApp.fetch(urlToken, tokenOptions);
  Logger.log(refreshTwitterXTokensRequest);
  const refreshTwitterXTokensResponse = JSON.parse(refreshTwitterXTokensRequest.getContentText());

  // ERROR HANDLING
  if (refreshTwitterXTokensRequest.getResponseCode() === 200 || refreshTwitterXTokensRequest.getResponseCode() === 201) {
    Logger.log('Successfully refreshed access tokens: ' + refreshTwitterXTokensRequest.getContentText());
    return refreshTwitterXTokensResponse;
  } else {
    Logger.log('ERROR: Unable to refresh access tokens. Response code: ' + refreshTwitterXTokensRequest.getResponseCode() + ' Response body: ' + refreshTwitterXTokensRequest.getContentText());
    return null;
  }
}
