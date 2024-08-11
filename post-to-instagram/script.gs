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

// POST TO INSTAGRAM
function postToInstagram(scriptProps, workbook, timesOfDay) {
  const ig = {
    appName: 'Instagram',
    accessToken: scriptProps.getProperty('igAccessToken'),
    expiration: parseInt(scriptProps.getProperty('igExpiresOn'), 10),
    businessAccountId: scriptProps.getProperty('igBusinessAccountId'),
    baseUrl: scriptProps.getProperty('igBaseUrl'),
    endpoint1: scriptProps.getProperty('igEndpoint1'),
    endpoint2: scriptProps.getProperty('igEndpoint2')
  };

  // CONFIRM ACTIVE TOKENS or REFRESH TOKENS
  // const instagramTokens = initializeInstagramTokens(scriptProps, ig);
  // if (!instagramTokens) {
    // Logger.log('ERROR: Did not receive tokens. Aborting execution.');
    // return;
  // };

  try {
    const columns = {
      status: columnIndex(ig.appName, 'status'),
      post_date: columnIndex(ig.appName, 'post_date'),
      time_of_day: columnIndex(ig.appName, 'time_of_day'),
      caption: columnIndex(ig.appName, 'caption'),
      is_carousel_item: columnIndex(ig.appName, 'is_carousel_item'),
      media_format: columnIndex(ig.appName, 'media_format'),
      media_type: columnIndex(ig.appName, 'media_type'),
      media_url: columnIndex(ig.appName, 'media_url'),
      location_id: columnIndex(ig.appName, 'location_id'),
      user_tags: columnIndex(ig.appName, 'user_tags'),
      product_tags: columnIndex(ig.appName, 'product_tags'),
      post_id: columnIndex(ig.appName, 'post_id'),
    };

    // SET ENDPOINTS
    const postContainer = `${ig.baseUrl}${ig.businessAccountId}${ig.endpoint1}`;
    const postSubmit = `${ig.baseUrl}${ig.businessAccountId}${ig.endpoint2}`;

    // GET ARRAY OF ALL POSTS WITHIN RANGE/TABLE
    const namedRange = workbook.getRangeByName(ig.appName);
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
      Logger.log('No Instagram posts scheduled for today at this time.');
      return; // Terminate if no posts are scheduled
    }

    // POST & RECORD DATA FOR EACH POST WITHIN FILTERED ARRAY
    filteredPosts.forEach((row, index) => {
      const mediaUrls = row[columns.media_url].split(',').map(url => url.trim());
      const caption = row[columns.caption];
      const location = row[columns.location_id];
      const mediaType = row[columns.media_type];

      // INITIATE creationId STRING VARIABLE
      let creationId;

      if (mediaType === 'CAROUSEL') {
        const mediaFormat = row[columns.media_format];
        creationId = createInstagramCarouselPost(mediaUrls, mediaFormat, caption, location, instagramTokens, postContainer);
      } else {
        creationId = createInstagramSinglePost(mediaUrls[0], false, mediaType, caption, location, instagramTokens, postContainer);
      }

      // SET PAYLOAD DATA
      const publishPayload = {
        'creation_id': creationId,
        'access_token': instagramTokens,
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
      const publishResponse = JSON.parse(publishRequest.getContentText());
      Logger.log(publishResponse.id);

      // RECORD POST DATA TO SPREADSHEET
      const rowIndex = namedRange.getRow() + index;
      workbook.getSheetByName(ig.appName).getRange(rowIndex, columns.post_id + 1).setValue(publishResponse.id);
      workbook.getSheetByName(ig.appName).getRange(rowIndex, columns.status + 1).setValue('POSTED');
    });
  } catch (error) {
    Logger.log('An error occurred while running postToInstagram:', error.message);
    return error;
  }
}

function createInstagramSinglePost(mediaUrl, carouselItem, mediaType, caption, location, access_token, postContainer) {
  try {
    let userTags = [];

    // INITIATE PAYLOAD DATA
    let postPayload = {
      'access_token': access_token,
      'is_carousel_item': carouselItem,
    };

    // ADD POST TYPE PAYLOAD DATA
    if (mediaType !== 'IMAGE') {
      // REELS: media_type, video_url, caption, share_to_feed, collaborators, cover_url, audio_name, user_tags, location_id, thumb_offset
      const cover = '';
      const shareToFeed = '';
      const thumbOffset = '';
      const audio = '';
      const collaborators = '';

      postPayload = {
        ...postPayload,
        'media_type': mediaType,
        'video_url': mediaUrl,
        ...(!carouselItem && caption && { 'caption': caption }),
        ...(!carouselItem && location && { 'location_id': location }),
        ...(!carouselItem && userTags.length > 0 && { 'user_tags': userTags }),
        ...(!carouselItem && cover && { 'cover_url': cover }),
        ...(!carouselItem && thumbOffset && { 'thumb_offset': thumbOffset }),
        ...(!carouselItem && audio && { 'audio_name': audio }),
        ...(!carouselItem && collaborators.length > 0 && { 'collaborators': collaborators }),
        ...(!carouselItem && shareToFeed && { 'share_to_feed': shareToFeed }),
      };
    } else {
      // IMAGE: image_url, caption, location_id, user_tags, product_tags
      postPayload = {
        ...postPayload,
        'image_url': mediaUrl,
        ...(!carouselItem && caption && { 'caption': caption }),
        ...(!carouselItem && location && { 'location_id': location }),
        ...(!carouselItem && userTags.length > 0 && { 'user_tags': userTags }),
      };
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
    Logger.log('An error occurred while running createInstagramSinglePost:', error.message);
    throw error;
  }
}

function createInstagramCarouselPost(mediaUrls, mediaFormat, caption, location, access_token, postContainer) {
  try {

    // INITIATE childern ARRAY VARIABLE
    let children = [];

    // CREATE EACH CAROUSEL ITEM & RETURN creationId TO ARRAY
    mediaUrls.forEach( (url) => {
      const creationId = createInstagramSinglePost(url, true, mediaFormat, '', location, access_token, postContainer);
      children.push(creationId);
    });

    Logger.log(children);

    const collaborators = '';
    const shareToFeed = '';
    const productTags = '';
    
    // SET PAYLOAD DATA
    // CAROUSEL: media_type, caption, share_to_feed, collaborators, location_id, product_tags, children, access_token
    const postPayload = {
      'media_type': 'CAROUSEL',
      'caption': caption,
      'children': children,
      'access_token': access_token,
        ...(collaborators.length > 0 && { 'collaborators': collaborators }),
        ...(shareToFeed && { 'share_to_feed': shareToFeed }),
        ...(location && { 'location_id': location }),
        ...(productTags.length > 0 && { 'product_tags': productTags }),
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
    Logger.log('An error occurred while running createInstagramCarouselPost:', error.message);
    throw error;
  }
}

// INITIALIZES OR UPDATES TOKENS
function initializeInstagramTokens(scriptProps, ig) {
  let access_token = ig.accessToken;
  let expires_on = new Date(ig.expiration);
  const expires_minus_1 = new Date(expires_on.getTime() - (24 * 60 * 60 *1000));

  // CONFIRM IF TODAY IS DAY BEFORE TOKEN EXPIRATION DATE
  if (new Date() > expires_minus_1 && new Date() < expires_on) {

    // REFRESH TOKENS (SEE ASSOCIATED FUNCTION)
    const refreshResponse = refreshInstagramTokens(access_token);

    // UPDATE SCRIPT PROPERTIES WITH REFRESHED TOKENS
    if (refreshResponse) {
      access_token = refreshResponse.access_token; // Get new access token
      const now = new Date(); // Get datetime of token refresh
      expires_on = now.getTime() + (refreshResponse.expires_in * 1000); // Set new token expiration date
      const expires_on_milliseconds = expires_on.toFixed(0); // Convert new token expiration date into milliseconds for storage
      scriptProps.setProperty('igAccessToken', access_token);
      scriptProps.setProperty('igExpiresOn', expires_on_milliseconds);
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
function refreshInstagramTokens(access_token) {
  const tokenPayload = {
    'grant_type': 'ig_refresh_token',
    'access_token': access_token
  };
  
  const tokenOptions = {
    method: 'GET',
    contentType: 'application/json',
    payload: tokenPayload,
    muteHttpExceptions: true
  };

  const urlRefreshToken = 'https://graph.instagram.com/refresh_access_token';

  const refreshInstagramTokensRequest = UrlFetchApp.fetch(urlRefreshToken + '?grant_type=ig_refresh_token&access_token=' + encodeURIComponent(access_token), tokenOptions);
  Logger.log(refreshInstagramTokensRequest);
  const refreshInstagramTokensResponse = JSON.parse(refreshInstagramTokensRequest.getContentText());

  if (refreshInstagramTokensRequest.getResponseCode() === 200) {
    Logger.log('Successfully refreshed Instagram tokens: ' + refreshInstagramTokensRequest.getContentText());
    return refreshInstagramTokensResponse;
  } else {
    Logger.log('ERROR: Unable to refresh Instagram tokens. Response code: ' + refreshInstagramTokensRequest.getResponseCode() + ' Response body: ' + refreshInstagramTokensRequest.getContentText());
    return null;
  }
}
