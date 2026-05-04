function inspectActivity() {
  const tokenResponse = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
    method: 'post',
    payload: {
      client_id: STRAVA_CLIENT_ID,
      client_secret: STRAVA_CLIENT_SECRET,
      refresh_token: STRAVA_REFRESH_TOKEN,
      grant_type: 'refresh_token'
    }
  });
  const accessToken = JSON.parse(tokenResponse.getContentText()).access_token;

  const activityId = '17872252619';
  const response = UrlFetchApp.fetch(
    `https://www.strava.com/api/v3/activities/${activityId}`,
    { headers: { 'Authorization': `Bearer ${accessToken}` } }
  );
  const activity = JSON.parse(response.getContentText());
  console.log(JSON.stringify(activity, null, 2));
}