var graph = require('@microsoft/microsoft-graph-client');

function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done) => {
      done(null, accessToken.accessToken);
    }
  });

  return client;
}

export async function getUserDetails(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const user = await client.api('/me').get();
  return user;
}

export async function getEvents(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const events = await client
    .api('/me/events')
    .select('subject,organizer,start,end')
    .orderby('createdDateTime DESC')
    .get();

  return events;
}

export async function getWorkBooksFromDrive(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const workbooks = await client
    .api(`me/drive/root/microsoft.graph.search(q='.xlsx')`)
    .select('name,id')
    .orderby('createdDateTime DESC')
    .get();

  return workbooks.value;
}

export async function getWorkSheetsForAWorkBookFromDrive(accessToken, workbookID) {
  const client = getAuthenticatedClient(accessToken);

  const worksheets = await client
    .api(`me/drive/items/${workbookID}/workbook/worksheets?$expand=charts`)
    .select('name,id')
    .get();

  return worksheets.value;
}

export async function getChartID(accessToken, workbookID, worksheetID) {
  const client = getAuthenticatedClient(accessToken);

  const charts = await client
    .api(`me/drive/items/${workbookID}/workbook/worksheets/${worksheetID}/Charts`)
    .select('name,id')
    .get();

  return charts.value;
}

export async function getChartImage(accessToken, workbookID, worksheetID, chartID) {
  const client = getAuthenticatedClient(accessToken);

  const chartImage = await client
    .api(`me/drive/items/${workbookID}/workbook/worksheets/${worksheetID}/Charts/${chartID}/Image`)
    .get();

  return chartImage.value;
}