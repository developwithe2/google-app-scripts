
function postToLinkedIn(scriptProps, workbook, timesOfDay) {
  const li = {
    appName: 'LinkedIn',
    accessToken: scriptProps.getProperty('liAccessToken'),
    accessExpiration: scriptProps.getProperty('liAccessExpiresOn'),
    refreshToken: scriptProps.getProperty('liRefreshToken'),
    refreshExpiration: scriptProps.getProperty('liRefreshExpiresOn'),
    clientId: scriptProps.getProperty('liClientId'),
    clientSecret: scriptProps.getProperty('liClientSecret'),
    profileId: scriptProps.getProperty('liProfileId'),
    baseUrl: scriptProps.getProperty('liBaseUrl'),
    endpointPosts: scriptProps.getProperty('liEndpointPosts'),
    endpointToken: scriptProps.getProperty('liEndpointToken'),
  };

  // CONFIRM ACTIVE TOKENS or REFRESH TOKENS
  const linkedinTokens = initializeLinkedInTokens(scriptProps, li);
  if (!linkedinTokens) {
    Logger.log('ERROR: Did not receive tokens. Aborting execution.');
    return;
  };

  try {
    const columns = {
      status: columnIndex(li.appName, 'status'),
      post_date: columnIndex(li.appName, 'post_date'),
      time_of_day: columnIndex(li.appName, 'time_of_day'),
      media_type: columnIndex(li.appName, 'media_type'),
      caption: columnIndex(li.appName, 'caption'),
      media_url: columnIndex(li.appName, 'media_url'),
      media_alt_tag: columnIndex(li.appName, 'media_alt_tag'),
      document_title: columnIndex(li.appName, 'document_title'),
      article_url: columnIndex(li.appName, 'article_url'),
      article_title: columnIndex(li.appName, 'article_title'),
      article_description: columnIndex(li.appName, 'article_description'),
      poll_responses: columnIndex(li.appName, 'poll_responses'),
      poll_duration: columnIndex(li.appName, 'poll_duration'),
      post_id: columnIndex(li.appName, 'post_id'),
    };

    // GET ARRAY OF ALL POSTS WITHIN RANGE/TABLE
    const namedRange = workbook.getRangeByName(li.appName);
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
      Logger.log('No LinkedIn posts scheduled for today at this time.');
      return; // Terminate run
    };

    // POST & RECORD DATA FOR EACH POST WITHIN FILTERED ARRAY
    filteredPosts.forEach((row, index) => {
      const mediaType = row[columns.media_type];
      const caption = row[columns.caption];
      const mediaUrl = row[columns.media_url];
      const mediaAltTag = row[columns.media_alt_tag];
      const documentTitle = row[columns.document_title];
      const articleUrl = row[columns.article_url];
      const articleTitle = row[columns.article_title];
      const articleDescription = row[columns.article_description];
      const pollResponses = row[columns.poll_responses].split(',').map(option => {
        return { 'text': option.trim() };
      });
      const pollDuration = row[columns.poll_duration];

      const postId = createLinkedInPost(li, linkedinTokens, mediaType, caption, mediaUrl, mediaAltTag, documentTitle, articleUrl, articleTitle, articleDescription, pollResponses, pollDuration);

      // RECORD POST DATA TO SPREADSHEET
      const rowIndex = namedRange.getRow() + index;
      workbook.getSheetByName(li.appName).getRange(rowIndex, columns.post_id + 1).setValue(postId);
      workbook.getSheetByName(li.appName).getRange(rowIndex, columns.status + 1).setValue('POSTED');
    });
  } catch (error) {
    Logger.log('An error occurred while running postToLinkedIn:', error.message);
  }
}

// CREATE POST
function createLinkedInPost(li, access_token, mediaType, caption, mediaUrl, mediaAltTag, documentTitle, articleUrl, articleTitle, articleDescription, pollResponses, pollDuration) {
  try {
    let mediaUrn = '';

    // UPLOAD MEDIA & OBTAIN A MEDIA URN
    if (mediaUrl) {
      mediaUrn = registerUploadToLinkedIn(li, access_token, mediaType, mediaUrl);
    }

    // INITIATE PAYLOAD
    let payload = {
      'author': `urn:li:person:${li.profileId}`, // Replace with your LinkedIn Profile ID, not Vanity Name
      'commentary': caption,
      'visibility': 'PUBLIC', // Options: PUBLIC, CONNECTIONS, LOGGED_IN, CONTAINER
      'distribution': {
        'feedDistribution': 'MAIN_FEED', // Options: MAIN_FEED, NONE
        'targetEntities': [], // For target marketing within LinkedIn, research later
        'thirdPartyDistributionChannels': [], // For posting to other apps
      },
      'lifecycleState': 'PUBLISHED',
      'isReshareDisabledByAuthor': false, // Boolean
    };

    // ADD POST TYPE PAYLOAD DATA
    if (mediaType === 'TEXT') {
      Logger.log('No additional payload needed.');
    } else if (mediaType === 'IMAGE' || 'VIDEO') {
      payload.content = {
        'media': {
          'id': mediaUrn, // URN of preloaded image or video
          'altText': mediaAltTag // Alt tag description
        }
      };
    } else if (mediaType === 'DOCUMENT') {
      // MAX 100MB - 300 PAGES; PPT, PPli, DOC, DOCX, PDF
      payload.content = {
        'media': {
          'id': mediaUrn, // URN of preloaded document
          'title': documentTitle // Document title
        }
      };
    } else if (mediaType === 'ARTICLE') {
      payload.content = {
        'article': {
          'source': articleUrl, // Url of article
          'thumbnail': mediaUrn, // URN of preloaded article image
          'title': articleTitle, // Article title
          'description': articleDescription // Article description
        }
      };
    } else if (mediaType === 'MULTIIMAGE') {
      // # OF IMAGES: MIN 2 - MAX 20
      payload.content = {
        'multiImage': {
          'images': [
            {
              'id': mediaUrn, // URN of preloaded image or video
              'altText': mediaAltTag // Alt tag description
            },
            {
              'id': mediaUrn, // URN of preloaded image or video
              'altText': mediaAltTag // Alt tag description
            }
          ],
        }
      };
    } else if (mediaType === 'POLL') {
      payload.content = {
        'poll': {
          'question': caption, // String
          'options': pollResponses, // Array of objects, MIN 2 - MAX 4
          'settings': {'duration': pollDuration} // Object, ONE_DAY, THREE_DAYS, SEVEN_DAYS, FOURTEEN_DAYS
        }
      };
    }

    // INITIATE OPTIONS
    let options = {
      'method': 'POST',
      'contentType': 'application/json',
      'muteHttpExceptions': true,
      'headers': {
        'Authorization': `Bearer ${access_token}`,
        'X-Restli-Protocol-Version': '2.0.0'
      },
      'payload': JSON.stringify(payload)
    };

    // SET ENDPOINT
    const urlPosts = `${li.baseUrl}${li.endpointPosts}`;

    // SEND & RECORD POST SUBMISSION
    const createLinkedInPostRequest = UrlFetchApp.fetch(urlPosts, options);
    Logger.log(createLinkedInPostRequest);
    let createLinkedInPostResponse = createLinkedInPostRequest.getAllHeaders();;
    let createLinkedInPostResponseId = createLinkedInPostResponse['x-restli-id'] || createLinkedInPostResponse['X-Restli-Id'];

    // ERROR HANDLING
    if (createLinkedInPostRequest.getResponseCode() === 200 || createLinkedInPostRequest.getResponseCode() === 201) {
      Logger.log('Successfully posted.');
      return createLinkedInPostResponseId;
    } else {
      Logger.log('ERROR: Unable to post. Response code: ' + createLinkedInPostRequest.getResponseCode() + '. Response body: ' + createLinkedInPostRequest.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log('An error occurred while running createLinkedInPost:', error.message);
  }
}

// UPLOAD MEDIA THROUGH UPLOAD URL
function registerUploadToLinkedIn(li, access_token, mediaType, mediaUrl) {
  try {
    // GET UPLOAD URL
    const uploadData = initializeUploadToLinkedIn(li, access_token, mediaType);
    const uploadUrl = uploadData.value.uploadUrl;
    let uploadMediaUrn = '';

    // GET MEDIA_URL BY DRIVE FILE ID
    // mediaUrl = DriveApp.getFileById('141VrJuWeSNNbVtOZ1hDGy82p6JEcah_a').getDownloadUrl();
    // Logger.log(mediaUrl);

    const media = UrlFetchApp.fetch(mediaUrl).getBlob();
    
    // INITIATE OPTIONS
    let options = {
      'method': 'PUT',
      'muteHttpExceptions': true,
      'headers': {
        'Authorization': `Bearer ${access_token}`,
        'X-Restli-Protocol-Version': '2.0.0'
      },
      'payload': media.getBytes()
    };
    if (mediaType === 'IMAGE') {
      uploadMediaUrn = uploadData.value.image;
      Logger.log(uploadMediaUrn);
      options = {
        ...options,
        'contentType': 'image/jpeg',
      }
    } else if (mediaType === 'VIDEO') {
      uploadMediaUrn = uploadData.value.video;
      Logger.log(uploadMediaUrn);
      options = {
        ...options,
        'contentType': 'application/octet-stream',
      }
    } else if (mediaType === 'DOCUMENT') {
      uploadMediaUrn = uploadData.value.document;
      Logger.log(uploadMediaUrn);
      options = {
        ...options,
        'contentType': 'application/octet-stream',
      }
    }

    // SEND & RECORD POST SUBMISSION, OUTPUTS: ???
    const registerUploadToLinkedInRequest = UrlFetchApp.fetch(uploadUrl, options);
    Logger.log(registerUploadToLinkedInRequest);
    const registerUploadToLinkedInResponse = registerUploadToLinkedInRequest.getAllHeaders();
    Logger.log(registerUploadToLinkedInResponse);

    // ERROR HANDLING
    if (registerUploadToLinkedInRequest.getResponseCode() === 200 || registerUploadToLinkedInRequest.getResponseCode() === 201) {
      Logger.log('Successfully registered upload.');
      return uploadMediaUrn;
    } else {
      Logger.log('ERROR: Unable to register upload. Response code: ' + registerUploadToLinkedInRequest.getResponseCode() + '. Response body: ' + registerUploadToLinkedInRequest.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log('An error occurred while running registerUploadToLinkedIn:', error.message);
  }
}

// REQUEST & OBTAIN A MEDIA UPLOAD URL
function initializeUploadToLinkedIn(li, access_token, mediaType) {
  try {
    // INITIATE mediaParam VARIABLE
    let mediaParam = 'images';

    // INITIATE UPLOAD REQUEST PAYLOAD
    let payload = {
      initializeUploadRequest: {
        owner: `urn:li:person:${li.profileId}`,
      }
    };
    if (mediaType === 'DOCUMENT') {
      mediaParam = 'documents';
    }
    if (mediaType === 'VIDEO') {
      mediaParam = 'videos';
      payload = {
        ...payload,
        fileSizeBytes: 0, // Integer
        uploadCaptions: false,
        uploadThumbnail: false
      }
    }

    // INITIATE OPTIONS
    let options = {
      'method': 'POST',
      'contentType': 'application/json',
      'muteHttpExceptions': true,
      'headers': {
        'Authorization': `Bearer ${access_token}`,
        'X-Restli-Protocol-Version': '2.0.0'
      },
      'payload': JSON.stringify(payload)
    };

    // SET ENDPOINT
    const urlInitUpload = `${li.baseUrl}/v2/${mediaParam}?action=initializeUpload`;

    // SEND & RECORD POST SUBMISSION, OUTPUTS: value.uploadUrl & value.image
    const initializeUploadToLinkedInRequest = UrlFetchApp.fetch(urlInitUpload, options);
    Logger.log(initializeUploadToLinkedInRequest);
    const initializeUploadToLinkedInResponse = JSON.parse(initializeUploadToLinkedInRequest.getContentText());
    Logger.log(initializeUploadToLinkedInResponse);

    // ERROR HANDLING
    if (initializeUploadToLinkedInRequest.getResponseCode() === 200 || initializeUploadToLinkedInRequest.getResponseCode() === 201) {
      Logger.log('Successfully initialized upload.');
      return initializeUploadToLinkedInResponse;
    } else {
      Logger.log('ERROR: Unable to initialize upload. Response code: ' + initializeUploadToLinkedInRequest.getResponseCode() + '. Response body: ' + initializeUploadToLinkedInRequest.getContentText());
      return false;
    }
  } catch(error) {
    Logger.log('An error occurred while running initializeUploadToLinkedIn:', error.message);
  }
}

// INITIALIZES OR UPDATES TOKENS
function initializeLinkedInTokens(scriptProps, li) {
  let access_token = li.accessToken;
  let refresh_token = li.refreshToken;
  let expires_on = new Date(li.accessExpiration);
  const expires_minus_1 = new Date(expires_on.getTime() - (24 * 60 * 60 *1000));

  // CONFIRM IF TODAY IS LATER THAN TOKEN EXPIRATION DATE
  if (new Date() > expires_minus_1 && new Date() < expires_on) {

    // REFRESH TOKENS (SEE ASSOCIATED FUNCTION)
    const refreshResponse = refreshLinkedInTokens(li, refresh_token);

    // UPDATE SCRIPT PROPERTIES WITH REFRESHED TOKENS
    if (refreshResponse) {
      access_token = refreshResponse.access_token;
      refresh_token = refreshResponse.refresh_token;
      const now = new Date(); // Get datetime of token refresh
      expires_on = now.getTime() + (refreshResponse.expires_in * 1000); // Set new token expiration date
      const expires_on_milliseconds = expires_on.toFixed(0); // Convert new token expiration date into milliseconds for storage
      scriptProps.setProperty('liAccessToken', access_token);
      // scriptProps.setProperty('liRefreshToken', refresh_token);
      scriptProps.setProperty('liAccessExpiresOn', expires_on_milliseconds);
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
function refreshLinkedInTokens(li, refresh_token) {
  const tokenPayload = {
    'grant_type': 'refresh_token',
    'refresh_token': refresh_token,
    'client_id': li.clientId,
    'client_secret': li.clientSecret
  };
  
  const tokenOptions = {
    method: 'POST',
    contentType: 'application/x-www-form-urlencoded',
    payload: Object.entries(tokenPayload).map(([key, value]) => encodeURIComponent(key) + '=' + encodeURIComponent(value)).join('&'),
    muteHttpExceptions: true
  };
  
  // SET ENDPOINT
  const urlRefreshToken = 'https://www.linkedin.com/oauth/v2/accessToken';

  // SEND TOKEN REFRESH REQUEST
  const refreshLinkedInTokensRequest = UrlFetchApp.fetch(urlRefreshToken, tokenOptions);
  Logger.log(refreshLinkedInTokensRequest);
  const refreshLinkedInTokensResponse = JSON.parse(refreshLinkedInTokensRequest.getContentText());
  
  if (refreshLinkedInTokensRequest.getResponseCode() === 200) {
    Logger.log('Successfully refreshed LinkedIn tokens: ' + refreshLinkedInTokensRequest.getContentText());
    return refreshLinkedInTokensResponse;
  } else {
    Logger.log('ERROR: Unable to refresh LinkedIn tokens. Response code: ' + refreshLinkedInTokensRequest.getResponseCode() + ' Response body: ' + refreshLinkedInTokensRequest.getContentText());
    return null;
  }
}
