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

// POST TO THREADS
function postToThreads(scriptProps, workbook, timesOfDay) {
  // SET th OBJECT WITH ALL SCRIPT PROPERTIES
  const th = {
    appName: 'Threads',
    accessToken: scriptProps.getProperty('thAccessToken'),
    expiration: parseInt(scriptProps.getProperty('thExpiresOn'), 10),
    userId: scriptProps.getProperty('thUserId'),
    baseUrl: scriptProps.getProperty('thBaseUrl'),
    endpoint1: scriptProps.getProperty('thEndpoint1'),
    endpoint2: scriptProps.getProperty('thEndpoint2')
  };

  // CONFIRM ACTIVE TOKENS or IF NOT, REFRESH TOKENS
  const threadsTokens = initializeThreadsTokens(scriptProps, th);
  if (!threadsTokens) {
    Logger.log('ERROR: Did not receive tokens. Aborting execution.');
    return;
  };

  try {
    // GET NECESSARY COLUMN INDEX #'S
    const columns = {
      status: columnIndex(th.appName, 'status'),
      post_date: columnIndex(th.appName, 'post_date'),
      time_of_day: columnIndex(th.appName, 'time_of_day'),
      caption: columnIndex(th.appName, 'caption'),
      is_carousel_item: columnIndex(th.appName, 'is_carousel_item'),
      media_format: columnIndex(th.appName, 'media_format'),
      media_type: columnIndex(th.appName, 'media_type'),
      media_url: columnIndex(th.appName, 'media_url'),
      post_id: columnIndex(th.appName, 'post_id'),
    };

    // SET ENDPOINTS
    const postContainer = `${th.baseUrl}${th.userId}${th.endpoint1}`;
    const postSubmit = `${th.baseUrl}${th.userId}${th.endpoint2}`;

    // GET ARRAY OF ALL POSTS WITHIN RANGE/TABLE
    const namedRange = workbook.getRangeByName(th.appName);
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

    // IF NO POSTS MATCH PARAMETERS, LOG & TERMINATE
    if (filteredPosts.length === 0) {
      Logger.log('No Threads posts scheduled for today at this time.');
      return;
    }

    // POST & RECORD EACH POST WITHIN FILTERED ARRAY
    filteredPosts.forEach((row, index) => {
      const mediaUrls = row[columns.media_url].split(',').map(url => url.trim());
      const caption = row[columns.caption];
      const mediaType = row[columns.media_type];

      // INITIATE creationId STRING VARIABLE
      let creationId;

      // DETERMINE TYPE OF POST
      if (mediaType === 'CAROUSEL') {
        const mediaFormat = row[columns.media_format];
        creationId = createThreadsCarouselPost(mediaUrls, mediaFormat, caption, threadsTokens, postContainer);
      } else {
        creationId = createThreadsSinglePost(mediaUrls[0], false, mediaType, caption, threadsTokens, postContainer);
      }

      Logger.log(creationId);

      // SET PAYLOAD DATA
      const publishPayload = {
        'creation_id': creationId,
        'access_token': threadsTokens,
      };

      // SET OPTIONS DATA
      const publishOptions = {
        'method': 'POST',
        'contentType': 'application/json',
        'payload': JSON.stringify(publishPayload),
        'muteHttpExceptions': true,
      };

      // SEND & RECORD POST SUBMISSION
      const publishRequest = UrlFetchApp.fetch(postSubmit, publishOptions);
      Logger.log(publishRequest);
      const publishResponse = JSON.parse(publishRequest.getContentText());
      Logger.log(publishResponse.id);

      // RECORD POST DATA TO SPREADSHEET
      const rowIndex = namedRange.getRow() + index;
      workbook.getSheetByName(th.appName).getRange(rowIndex, columns.post_id + 1).setValue(publishResponse.id);
      workbook.getSheetByName(th.appName).getRange(rowIndex, columns.status + 1).setValue('POSTED');
    });
  } catch (error) {
    Logger.log(`An error occurred while running postToThreads: ${error.message}`);
    return error;
  }
}

function createThreadsSinglePost(mediaUrl, carouselItem, mediaType, caption, access_token, postContainer) {
  try {

    // INITIATE PAYLOAD DATA
    let postPayload = {
      'access_token': access_token,
      'media_type': mediaType,
    };

    // ADD POST TYPE PAYLOAD DATA
    if (mediaType === 'TEXT') {
      postPayload = {
        ...postPayload,
        'text': caption,
      }
    } else if (mediaType === 'IMAGE') {
      postPayload = {
        ...postPayload,
        'image_url': mediaUrl,
        'is_carousel_item': carouselItem,
        ...(!carouselItem && caption && { 'text': caption }),
      };
    } else if (mediaType === 'VIDEO') {
      postPayload = {
        ...postPayload,
        'video_url': mediaUrl,
        'is_carousel_item': carouselItem,
        ...(!carouselItem && caption && { 'text': caption }),
      };
    } else {
      Logger.log(`Incorrect media type chosen: ${mediaType}`);
      return;
    }

    // SET OPTIONS DATA
    const postOptions = {
      'method': 'POST',
      'contentType': 'application/json',
      'payload': JSON.stringify(postPayload),
      'muteHttpExceptions': true,
    };

    // SEND & RECORD POST REQUEST
    const publishSinglePostRequest = UrlFetchApp.fetch(postContainer, postOptions);
    Logger.log(publishSinglePostRequest);
    const creation = JSON.parse(publishSinglePostRequest.getContentText());
    Logger.log(creation.id);
    return creation.id;
  } catch (error) {
    Logger.log('An error occurred while running createThreadsSinglePost:', error.message);
    throw error;
  }
}

function createThreadsCarouselPost(mediaUrls, mediaFormat, caption, access_token, postContainer) {
  try {
    // INITIATE childern ARRAY VARIABLE
    let children = [];

    // CREATE EACH CAROUSEL ITEM & RETURN creationId TO ARRAY
    mediaUrls.forEach( (url) => {
      const creationId = createThreadsSinglePost(url, true, mediaFormat, '', access_token, postContainer);
      children.push(creationId);
    });

    Logger.log(children);
    
    // SET PAYLOAD DATA
    // CAROUSEL: media_type, caption, children, access_token
    const postPayload = {
      'media_type': 'CAROUSEL',
      'text': caption,
      'children': children,
      'access_token': access_token,
    };

    // SET OPTIONS DATA
    const postOptions = {
      'method': 'POST',
      'contentType': 'application/json',
      'payload': JSON.stringify(postPayload),
      'muteHttpExceptions': true,
    };

    // SEND & RECORD POST REQUEST
    const publishCarouselPostRequest = UrlFetchApp.fetch(postContainer, postOptions);
    Logger.log(publishCarouselPostRequest);
    const creation = JSON.parse(publishCarouselPostRequest.getContentText());
    Logger.log(creation.id);
    return creation.id;
  } catch (error) {
    Logger.log('An error occurred while running createThreadsCarouselPost:', error.message);
    throw error;
  }
}

// INITIALIZES OR UPDATES TOKENS
function initializeThreadsTokens(scriptProps, th) {
  // INITIALIZE & SET VARIABLES
  let access_token = th.accessToken;
  let expires_on = new Date(th.expiration);
  const expires_minus_1 = new Date(expires_on.getTime() - (24 * 60 * 60 *1000));

  // CONFIRM IF TODAY IS DAY BEFORE TOKEN EXPIRATION DATE
  if (new Date() > expires_minus_1 && new Date() < expires_on) {

    // REFRESH TOKENS (SEE ASSOCIATED FUNCTION)
    const refreshResponse = refreshThreadsTokens(access_token);

    // UPDATE SCRIPT PROPERTIES WITH REFRESHED DATA
    if (refreshResponse) {
      access_token = refreshResponse.access_token; // Get new access token
      const now = new Date(); // Get datetime of token refresh
      expires_on = now.getTime() + (refreshResponse.expires_in * 1000); // Set new token expiration date
      const expires_on_milliseconds = expires_on.toFixed(0); // Convert new token expiration date into milliseconds for storage
      scriptProps.setProperty('thAccessToken', access_token);
      scriptProps.setProperty('thExpiresOn', expires_on_milliseconds);
    } else {
      return null;
    }
  } else {
    Logger.log('All tokens are currently active.');
  }
  // RETURN EXISTING TOKENS OR REFRESHED TOKENS
  return access_token;
}

// REFRESH TOKEN
function refreshThreadsTokens(access_token) {
  const tokenPayload = {
    'grant_type': 'th_refresh_token',
    'access_token': access_token
  };
  
  const tokenOptions = {
    method: 'GET',
    contentType: 'application/json',
    payload: tokenPayload,
    muteHttpExceptions: true
  };

  const urlRefreshToken = 'https://graph.threads.net/refresh_access_token';

  const refreshThreadsTokensRequest = UrlFetchApp.fetch(urlRefreshToken + '?grant_type=th_refresh_token&access_token=' + encodeURIComponent(access_token), tokenOptions);
  Logger.log(refreshThreadsTokensRequest);
  const refreshThreadsTokensResponse = JSON.parse(refreshThreadsTokensRequest.getContentText());

  if (refreshThreadsTokensRequest.getResponseCode() === 200) {
    Logger.log('Successfully refreshed Instagram tokens: ' + refreshThreadsTokensRequest.getContentText());
    return refreshThreadsTokensResponse;
  } else {
    Logger.log('ERROR: Unable to refresh Instagram tokens. Response code: ' + refreshThreadsTokensRequest.getResponseCode() + ' Response body: ' + refreshThreadsTokensRequest.getContentText());
    return null;
  }
}
