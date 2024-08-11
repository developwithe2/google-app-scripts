# Post to Social Media from Google Sheets

## Overview

A Google App Script project that schedules & publishes social media posts to the following platforms:
- [Instagram](https://developers.facebook.com/docs/instagram-basic-display-api/overview)
- [Threads](https://developers.facebook.com/docs/threads/get-started)
- [Twitter X](https://developer.x.com/en/docs/twitter-api/tweets/manage-tweets/introduction)
- [LinkedIn](https://learn.microsoft.com/en-us/linkedin/marketing/quick-start?view=li-lms-2024-07)

This project is a centralized social media management system that accomplishes the following solutions:
- Manage, schedule & publish posts across multiple accounts & platforms
- Schedule posts as far in advance as desired
- Synchronize post campaigns across multiple platforms
- Conditional scripts run based on selected post media type to ensure appropriate payload data & endpoints
- Customize `time_of_day` to match preferred daily posting timeframes
- Utilizes a familiar, cross-generational spreadsheet interface
- Auto-refresh access tokens based on token expiration timelines

Additional solutions currently being developed:
- Posting schedule synchronization with Google Calendars
- Text notifications using Google Voice via GMail

## Basic Information

### Setting Up of the Google Sheet

Within the workbook, each app has a designated sheet. Each table/range is structured based on specific parameter keys of the specific API.

This project utilizes named ranges. Each sheet will consist of two (2) named ranges:
- headers: single row of all range headers
- posts: multiple rows of all post information

| post_date | time_of_day | media_type            | caption                 | media_url       | is_carousel_item | .... |
| :-------- | :---------- | :-------------------- | :---------------------- | :-------------- | :--------------- | :--- |
| 9/1/2024  | Morning     | TEXT/IMAGE/VIDEO/etc. | Post caption goes here. | URL to image... | YES/NO           | etc. |

**Table 1:** Example of sheet configuration (shortened). In this version, table headers are labelled to match the parameter keys set by the API documentation and are utilized within the code to index columns.
