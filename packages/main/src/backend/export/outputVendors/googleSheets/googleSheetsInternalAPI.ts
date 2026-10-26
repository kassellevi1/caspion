import { google, type Auth, type drive_v3 as driveV3 } from 'googleapis';

export type OAuth2Client = Auth.OAuth2Client;
export type Credentials = Auth.Credentials;

const sheets = google.sheets({ version: 'v4' });
const drive = google.drive({ version: 'v3' });

/**
 * Prints the names and majors of students in a sample spreadsheet:
 * @see https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
export async function appendToSpreadsheet(
  spreadsheetId: string,
  range: string,
  values: unknown[][],
  auth: OAuth2Client,
) {
  const requestBody = {
    values,
  };

  const result = await sheets.spreadsheets.values.append({
    auth,
    spreadsheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody,
  });
  return result;
}

export interface SheetRow {
  rowIndex: number; // 0-based, including the header row
  hash: string;
  status: string;
}

export async function getExistingRows(
  spreadsheetId: string,
  sheetName: string,
  auth: OAuth2Client,
  hashColumn = 'G',
  statusColumn = 'M',
): Promise<SheetRow[]> {
  const result = await sheets.spreadsheets.values.batchGet({
    auth,
    spreadsheetId,
    ranges: [`${sheetName}!${hashColumn}:${hashColumn}`, `${sheetName}!${statusColumn}:${statusColumn}`],
  });
  // One entry per row; empty cells come back as empty arrays, so don't flatten (it would shift the row indexes)
  const [hashes = [], statuses = []] = (result.data.valueRanges ?? []).map(({ values }) =>
    (values ?? []).map((row) => String(row[0] ?? '')),
  );
  return hashes.map((hash, rowIndex) => ({ rowIndex, hash, status: statuses[rowIndex] ?? '' }));
}

export async function getColumnValues(
  spreadsheetId: string,
  sheetName: string,
  column: string,
  auth: OAuth2Client,
): Promise<string[]> {
  const result = await sheets.spreadsheets.values.get({
    auth,
    spreadsheetId,
    range: `${sheetName}!${column}:${column}`,
  });
  return (result.data.values ?? []).map((row) => String(row[0] ?? ''));
}

export async function addSheet(spreadsheetId: string, sheetTitle: string, auth: OAuth2Client) {
  return sheets.spreadsheets.batchUpdate({
    auth,
    spreadsheetId,
    requestBody: { requests: [{ addSheet: { properties: { title: sheetTitle } } }] },
  });
}

export async function updateCells(
  spreadsheetId: string,
  data: { range: string; values: unknown[][] }[],
  auth: OAuth2Client,
) {
  return sheets.spreadsheets.values.batchUpdate({
    auth,
    spreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data },
  });
}

export async function deleteRows(spreadsheetId: string, sheetId: number, rowIndexes: number[], auth: OAuth2Client) {
  // Delete from the bottom up so earlier deletions don't shift the indexes of later ones
  const requests = [...rowIndexes]
    .sort((a, b) => b - a)
    .map((rowIndex) => ({
      deleteDimension: {
        range: { sheetId, dimension: 'ROWS', startIndex: rowIndex, endIndex: rowIndex + 1 },
      },
    }));
  return sheets.spreadsheets.batchUpdate({ auth, spreadsheetId, requestBody: { requests } });
}

export type Spreadsheet = Pick<driveV3.Schema$File, 'id' | 'name'>;
export const getAllSpreadsheets = async (auth: OAuth2Client) => {
  const response = await drive.files.list({
    q: 'mimeType="application/vnd.google-apps.spreadsheet"',
    auth,
  });

  return response.data.files as Spreadsheet[];
};

export async function getSheet(spreadsheetId: string, sheetName: string, auth: OAuth2Client) {
  const spreadsheetResponse = await sheets.spreadsheets.get({
    auth,
    spreadsheetId,
  });
  const sheet = spreadsheetResponse.data.sheets?.find(({ properties }) => properties?.title === sheetName);

  return sheet;
}

export async function createSpreadsheet(
  spreadsheetTitle: string,
  sheetTitle: string,
  auth: OAuth2Client,
): Promise<string> {
  const spreadsheetResponse = await sheets.spreadsheets.create({
    auth,
    requestBody: {
      properties: {
        title: spreadsheetTitle,
      },
      sheets: [{ properties: { title: sheetTitle } }],
    },
  });

  const { spreadsheetId } = spreadsheetResponse.data;
  return spreadsheetId!;
}
