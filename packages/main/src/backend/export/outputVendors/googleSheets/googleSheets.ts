import {
  OutputVendorName,
  type AccountBalance,
  type EnrichedTransaction,
  type ExportTransactionsFunction,
  type OutputVendor,
} from '@/backend/commonTypes';
import { EventNames, ExporterEvent, type EventPublisher } from '@/backend/eventEmitters/EventEmitter';
import { filterExistedHashes, unifyHash } from '@/backend/transactions/transactions';
import { type Auth } from 'googleapis';
import moment from 'moment/moment';
import { createClient } from './googleAuth';
import * as googleSheets from './googleSheetsInternalAPI';
import { logAppEvent } from '/@/logging/operationLogger';

const GOOGLE_SHEETS_DATE_FORMAT = 'YYYY-MM-DD';
const DEFAULT_SHEET_NAME = '_caspion';
const COLUMN_HEADERS = [
  'תאריך',
  'סכום',
  'תיאור',
  'תיאור נוסף',
  'קטגוריה',
  'מספר חשבון',
  'hash - לא לגעת',
  'מטבע',
  'סכום מקורי',
  'ת. חיוב',
  'מזהה',
  'סוג',
  'סטטוס',
];
const PROCESSED_DATE_COLUMN = 'J';
const STATUS_COLUMN = 'M';
const TRANSACTION_STATUS_PENDING = 'pending';
const TRANSACTION_STATUS_COMPLETED = 'completed';

const BALANCE_SHEET_NAME = 'balance';
const BALANCE_COLUMN_HEADERS = ['תאריך', 'חברה', 'מספר חשבון', 'יתרה', 'עודכן', 'מפתח - לא לגעת'];
const BALANCE_KEY_COLUMN = 'F';
const BALANCE_UPDATED_AT_FORMAT = 'YYYY-MM-DD HH:mm';

// The hash is `${date}_${chargedAmount}_${description}_${memo}_${companyId}_${accountNumber}` (see calculateTransactionHash)
const getHashDate = (hash: string) => moment(hash.split('_')[0], moment.ISO_8601, true);
const getHashAccount = (hash: string) => hash.split('_').slice(-2).join('_');

/**
 * Decides what to do with rows that were exported while their transaction was still pending:
 * - The same transaction is now completed: update the row's status and processed date.
 * - The transaction no longer appears in a fresh scrape of the same account and date range
 *   (it was captured with a different amount/description, or canceled): delete the row.
 */
export function getPendingRowChanges(
  sheetRows: googleSheets.SheetRow[],
  scrapedTransactions: EnrichedTransaction[],
  startDate: Date,
) {
  const scrapedByHash = new Map(scrapedTransactions.map((transaction) => [unifyHash(transaction.hash), transaction]));
  const scrapedAccounts = new Set(scrapedTransactions.map(({ hash }) => getHashAccount(unifyHash(hash))));
  // A day of margin so time zone differences around the start date never delete a row that simply wasn't scraped
  const safeStartDate = moment(startDate).add(1, 'day');

  const rowsToComplete: { rowIndex: number; transaction: EnrichedTransaction }[] = [];
  const rowIndexesToDelete: number[] = [];

  sheetRows
    .filter(({ status }) => status === TRANSACTION_STATUS_PENDING)
    .forEach(({ rowIndex, hash }) => {
      const unifiedHash = unifyHash(hash);
      const scrapedTransaction = scrapedByHash.get(unifiedHash);
      if (scrapedTransaction) {
        if (scrapedTransaction.status === TRANSACTION_STATUS_COMPLETED) {
          rowsToComplete.push({ rowIndex, transaction: scrapedTransaction });
        }
        return;
      }
      const hashDate = getHashDate(unifiedHash);
      if (
        scrapedAccounts.has(getHashAccount(unifiedHash)) &&
        hashDate.isValid() &&
        hashDate.isSameOrAfter(safeStartDate)
      ) {
        rowIndexesToDelete.push(rowIndex);
      }
    });

  return { rowsToComplete, rowIndexesToDelete };
}

/**
 * Keeps one row per account per day in the balance sheet: the first run of the day appends a row,
 * later runs on the same day overwrite it with the latest balance.
 */
export function getBalanceRowChanges(existingKeys: string[], balances: AccountBalance[], now: Date) {
  const date = moment(now).format(GOOGLE_SHEETS_DATE_FORMAT);
  const updatedAt = moment(now).format(BALANCE_UPDATED_AT_FORMAT);

  const rowsToUpdate: { rowIndex: number; values: unknown[] }[] = [];
  const rowsToAppend: unknown[][] = [];

  balances.forEach(({ companyId, accountNumber, balance }) => {
    const key = `${date}_${companyId}_${accountNumber}`;
    const values = [date, companyId, accountNumber, balance, updatedAt, key];
    const rowIndex = existingKeys.indexOf(key);
    if (rowIndex === -1) {
      rowsToAppend.push(values);
    } else {
      rowsToUpdate.push({ rowIndex, values });
    }
  });

  return { rowsToUpdate, rowsToAppend };
}

async function exportBalances(spreadsheetId: string, balances: AccountBalance[], auth: googleSheets.OAuth2Client) {
  if (balances.length === 0) {
    return;
  }

  if (!(await googleSheets.getSheet(spreadsheetId, BALANCE_SHEET_NAME, auth))) {
    logAppEvent('GOOGLE_SHEETS_CREATING_BALANCE_SHEET');
    await googleSheets.addSheet(spreadsheetId, BALANCE_SHEET_NAME, auth);
    await googleSheets.appendToSpreadsheet(spreadsheetId, `${BALANCE_SHEET_NAME}!A:A`, [BALANCE_COLUMN_HEADERS], auth);
  }

  const existingKeys = await googleSheets.getColumnValues(spreadsheetId, BALANCE_SHEET_NAME, BALANCE_KEY_COLUMN, auth);
  const { rowsToUpdate, rowsToAppend } = getBalanceRowChanges(existingKeys, balances, new Date());

  logAppEvent('GOOGLE_SHEETS_EXPORTING_BALANCES', { updated: rowsToUpdate.length, appended: rowsToAppend.length });
  if (rowsToUpdate.length) {
    await googleSheets.updateCells(
      spreadsheetId,
      rowsToUpdate.map(({ rowIndex, values }) => ({
        range: `${BALANCE_SHEET_NAME}!A${rowIndex + 1}`,
        values: [values],
      })),
      auth,
    );
  }
  if (rowsToAppend.length) {
    await googleSheets.appendToSpreadsheet(spreadsheetId, `${BALANCE_SHEET_NAME}!A:A`, rowsToAppend, auth);
  }
}

const createTransactionsInGoogleSheets: ExportTransactionsFunction = async (
  { transactionsToCreate: transactions, balances, startDate, outputVendorsConfig },
  eventPublisher,
) => {
  const startTime = Date.now();
  const { spreadsheetId, credentials } = outputVendorsConfig.googleSheets!.options;

  logAppEvent('GOOGLE_SHEETS_EXPORT_START', {
    hasSpreadsheetId: !!spreadsheetId,
    hasCredentials: !!credentials,
    transactionsCount: transactions.length,
  });

  if (!credentials) {
    logAppEvent('GOOGLE_SHEETS_ERROR', { error: 'No credentials provided' });
    throw new Error("You must set the 'credentials'");
  }

  const oAuthClient = createClient(credentials);

  logAppEvent('GOOGLE_SHEETS_CHECKING_SHEET');
  const sheet = await googleSheets.getSheet(spreadsheetId, DEFAULT_SHEET_NAME, oAuthClient);
  if (!sheet) {
    logAppEvent('GOOGLE_SHEETS_ERROR', { error: `Sheet ${DEFAULT_SHEET_NAME} not found` });
    throw new Error(`There is no sheet called ${DEFAULT_SHEET_NAME} in the spreadsheet`);
  }

  logAppEvent('GOOGLE_SHEETS_FETCHING_EXISTING_HASHES');
  const existingRows = await googleSheets.getExistingRows(spreadsheetId, DEFAULT_SHEET_NAME, oAuthClient);
  const transactionsToCreate = filterExistedHashes(
    transactions,
    existingRows.map(({ hash }) => hash),
  );

  logAppEvent('GOOGLE_SHEETS_DUPLICATE_CHECK_COMPLETE', {
    totalTransactions: transactions.length,
    alreadyExist: transactions.length - transactionsToCreate.length,
    toCreate: transactionsToCreate.length,
  });

  const { rowsToComplete, rowIndexesToDelete } = getPendingRowChanges(existingRows, transactions, startDate);

  // Update before deleting, since deleting rows shifts the row numbers below them
  if (rowsToComplete.length) {
    logAppEvent('GOOGLE_SHEETS_COMPLETING_PENDING', { rowCount: rowsToComplete.length });
    await googleSheets.updateCells(
      spreadsheetId,
      rowsToComplete.flatMap(({ rowIndex, transaction }) => [
        {
          range: `${DEFAULT_SHEET_NAME}!${PROCESSED_DATE_COLUMN}${rowIndex + 1}`,
          values: [[moment(transaction.processedDate).format(GOOGLE_SHEETS_DATE_FORMAT)]],
        },
        { range: `${DEFAULT_SHEET_NAME}!${STATUS_COLUMN}${rowIndex + 1}`, values: [[transaction.status]] },
      ]),
      oAuthClient,
    );
  }

  if (rowIndexesToDelete.length) {
    logAppEvent('GOOGLE_SHEETS_DELETING_STALE_PENDING', { rowCount: rowIndexesToDelete.length });
    await googleSheets.deleteRows(spreadsheetId, sheet.properties!.sheetId!, rowIndexesToDelete, oAuthClient);
  }

  if (transactionsToCreate.length === 0) {
    await emitProgressEvent(eventPublisher, transactions, 'All transactions already exist in google sheets');
  } else {
    await emitProgressEvent(
      eventPublisher,
      transactions,
      `Creating ${transactionsToCreate.length} transactions in google sheets`,
    );

    const transactionsInSheetsFormat = transactionsToCreate.map((transaction) => [
      moment(transaction.date).format(GOOGLE_SHEETS_DATE_FORMAT),
      transaction.chargedAmount,
      transaction.description,
      transaction.memo,
      transaction.category,
      transaction.accountNumber,
      transaction.hash,
      transaction.originalCurrency,
      transaction.originalAmount,
      moment(transaction.processedDate).format(GOOGLE_SHEETS_DATE_FORMAT),
      transaction.identifier,
      transaction.type,
      transaction.status,
    ]);

    logAppEvent('GOOGLE_SHEETS_APPENDING_DATA', { rowCount: transactionsInSheetsFormat.length });
    await googleSheets.appendToSpreadsheet(
      spreadsheetId,
      `${DEFAULT_SHEET_NAME}!A:A`,
      transactionsInSheetsFormat,
      oAuthClient,
    );
  }

  // After the transactions, so a balance sheet problem never blocks exporting them
  await exportBalances(spreadsheetId, balances, oAuthClient);

  logAppEvent('GOOGLE_SHEETS_SUCCESS', {
    createdCount: transactionsToCreate.length,
    duration: `${Date.now() - startTime}ms`,
  });
  return {
    exportedTransactionsNum: transactionsToCreate.length,
  };
};

async function emitProgressEvent(
  eventPublisher: EventPublisher,
  allTransactions: EnrichedTransaction[],
  message: string,
) {
  await eventPublisher.emit(
    EventNames.EXPORTER_PROGRESS,
    new ExporterEvent({
      message,
      exporterName: googleSheetsOutputVendor.name,
      allTransactions,
    }),
  );
}

export async function createSpreadsheet(spreadsheetTitle: string, credentials: Auth.Credentials): Promise<string> {
  const auth = createClient(credentials);
  const spreadsheetId = await googleSheets.createSpreadsheet(spreadsheetTitle, DEFAULT_SHEET_NAME, auth);

  await googleSheets.appendToSpreadsheet(spreadsheetId, `${DEFAULT_SHEET_NAME}!A:A`, [COLUMN_HEADERS], auth);

  return spreadsheetId;
}

export const googleSheetsOutputVendor: OutputVendor = {
  name: OutputVendorName.GOOGLE_SHEETS,
  exportTransactions: createTransactionsInGoogleSheets,
  supportsPendingTransactions: true,
};
