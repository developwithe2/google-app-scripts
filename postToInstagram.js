// POST TO INSTAGRAM VIA GOOGLE SHEETS & GOOGLE APP SCRIPT

// INSTAGRAM GRAPH API FOR BUSINESS ACCOUNT
// Create a Facebook app through Meta Developers
// Assign app to a Facebook Business account
// Access token generated through Meta Business Manager for System User
// Retrieve Business Account ID through Facebook Graph API Explorer
  // GET graph.facebook.com/{api_version}/me/accounts?fields=name,id,access_token,instagram_business_account{id,username,profile_picture_url}

// BASE VARIABLES
const workbook = SpreadsheetApp.getActiveSpreadsheet();
// Index integer of columns in spreadsheet
const columns = {
  title: 0, // String
  post_date: 1, // Date
  time_of_day: 2, // String
  post_time: 3, // Time
  is_carousel_item: 4, // String, 'Yes' or 'No'
  media_type: 5, // String, Refer to API documentation for all types
  media_url: 6, // String array
  caption: 7, // String
}
// Set time-of-day phases & houe ranges
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
// Configure Intagram API
const igAccess = [SET_INTAGRAM_ACCESS_KEY];
const igBusinessAccountId = [SET_INTRAGRAM_ACCOUNT_ID];
const igBaseUrl = [SET_INSTAGRAM_BASE_URL];
const igEndpoint1 = [SET_INSTAGRAM_ENDPOINT1];
const igEndpoint2 = [SET_INSTAGRAM_ENDPOINT2];
const postContainer = `${igBaseUrl}${igBusinessAccountId}${igEndpoint1}`;
const postSubmit = `${igBaseUrl}${igBusinessAccountId}${igEndpoint2}`;

// MAIN FUNCTION
function postToInstagram() {
  try {
    const namedRange = workbook.getRangeByName('[ENTER_RANGE_NAME]'); // Set namedRange variable
    const postRows = namedRange.getValues(); // Get all values in named ranged

    // Setting date & time-of-day associated variables for filtering
    const now = new Date();
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const currentHour = now.getHours();
    const currentTimeOfDay = Object.keys(timesOfDay).find(period => {
      const [start, end] = timesOfDay[period];
      return currentHour >= start && currentHour < end;
    });

    // Filter post rows by date & time-of-day
    const filteredPosts = postRows.filter(row => {
      const postDate = Utilities.formatDate(new Date(row[columns.post_date]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const postTimeOfDay = row[columns.time_of_day];
      return postDate === today && postTimeOfDay === currentTimeOfDay;
    });
    if (filteredPosts.length === 0) {
      Logger.log('No posts scheduled for today at this time.');
      return; // Terminate script if no posts are scheduled for date & time-of-day
    };

    filteredPosts.forEach((row, index) => {
      const mediaUrls = row[columns.media_url].split(',').map(url => url.trim()); // Convert multi-values to array
      const caption = row[columns.caption];
      const location = row[columns.location_id];
      const mediaType = row[columns.media_type];

      let creationId;

      // Determine what media type is being posted
      if (mediaType === 'CAROUSEL') {
        creationId = createCarouselPost(mediaUrls, caption, location, mediaType);
      } else {
        creationId = createSinglePost(mediaUrls[0], caption, location, false, mediaType);
      };

      const publishData = {
        'creation_id': creationId,
        'access_token': igAccess,
      };
      const publishOptions = {
        'method': 'post',
        'payload': publishData,
        "muteHttpExceptions": true,
      };
      const publishResponse = UrlFetchApp.fetch(postSubmit, publishOptions);
      Logger.log(publishResponse.getContentText());

      const response = publishResponse.getContentText();
      const post = JSON.parse(response);

      // Record the post id back to spreadsheet in the post_id column
      const rowIndex = namedRange.getRow() + index;
      workbook.getSheetByName('[ENTER_SHEET_NAME]').getRange(rowIndex, columns.post_id + 1).setValue(post.id);
    });
  } catch (error) {
    Logger.log('Error running postToInstagram:', error.message);
    return error;
  } finally {
    Logger.log('postToInstagram finished running.');
  };
};

/**
 * Function to create a single post
 * @param {string} mediaUrl - URL of the media to post
 * @param {string} caption - Caption for the post
 * @param {string} location - Location ID for the post
 * @param {boolean} carouselItem - Flag indicating if the media is part of a carousel
 * @param {string} mediaType - Type of the media (IMAGE, VIDEO, etc.)
 * @return {string} - Creation ID of the post
 */
function createSinglePost(mediaUrl, caption, location, carouselItem, mediaType) {
  try {
    const userTags = [
      {
        username: 'developwithe2',
        x: 0.1,
        y: 0.1,
      }
    ];

    let postData = {
      'access_token': igAccess,
      'is_carousel_item': carouselItem,
    };

    if (mediaType !== 'IMAGE') {
      // REELS: media_type, video_url, caption, share_to_feed, collaborators, cover_url, audio_name, user_tags, location_id, thumb_offset, access_token
      const cover = '';
      const shareToFeed = '';
      const thumbOffset = '';
      const audio = '';
      const collaborators = '';

      postData = {
        ...postData,
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
      // IMAGE: image_url, caption, location_id, user_tags, product_tags, is_carousel_item, access_token
      postData = {
        ...postData,
        'image_url': mediaUrl,
        ...(!carouselItem && caption && { 'caption': caption }),
        ...(!carouselItem && location && { 'location_id': location }),
        ...(!carouselItem && userTags.length > 0 && { 'user_tags': userTags }),
      };
    };

    const postOptions = {
      'method': 'post',
      'payload': postData,
      "muteHttpExceptions": true,
    };
    const publishRequest = UrlFetchApp.fetch(postContainer, postOptions);
    Logger.log(publishRequest.getContentText());

    const container = publishRequest.getContentText();
    const creation = JSON.parse(container);
    return creation.id;
  } catch (error) {
    Logger.log('Error running createSinglePost:', error.message);
    throw error; // Rethrow error to propagate it
  } finally {
    Logger.log('createSinglePost finished running.');
  };
};

/**
 * Function to create a carousel post
 * @param {Array<string>} mediaUrls - Array of media URLs for the carousel
 * @param {string} caption - Caption for the post
 * @param {string} location - Location ID for the post
 * @param {string} mediaType - Type of the media
 * @return {string} - Creation ID of the post
 */
function createCarouselPost(mediaUrls, caption, location, mediaType) {
  try {
    const children = [];

    mediaUrls.forEach( (url) => {
      const creationId = createSinglePost(url, '', location, true, mediaType);
      children.push(creationId);
    });

    console.log(children);

    const collaborators = '';
    const shareToFeed = '';
    const productTags = '';
    
    // CAROUSEL: media_type, caption, share_to_feed, collaborators, location_id, product_tags, children, access_token
    const postData = {
      'media_type': 'CAROUSEL',
      'caption': caption,
      'children': children,
      'access_token': igAccess,
        ...(collaborators.length > 0 && { 'collaborators': collaborators }),
        ...(shareToFeed && { 'share_to_feed': shareToFeed }),
        ...(location && { 'location_id': location }),
        ...(productTags.length > 0 && { 'product_tags': productTags }),
    };

    const postOptions = {
      'method': 'post',
      'payload': postData,
      "muteHttpExceptions": true,
    };
    const publishRequest = UrlFetchApp.fetch(postContainer, postOptions);
    Logger.log(publishRequest.getContentText());

    const container = publishRequest.getContentText();
    const creation = JSON.parse(container);
    return creation.id;
  } catch (error) {
    Logger.log('Error running createCarouselPost:', error.message);
    throw error; // Rethrow error to propagate it
  } finally {
    Logger.log('createCarouselPost finished running.');
  };
};
