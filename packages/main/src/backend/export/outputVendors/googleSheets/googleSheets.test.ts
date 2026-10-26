import { type EnrichedTransaction } from '@/backend/commonTypes';
import { describe, expect, test } from 'vitest';
import { getPendingRowChanges } from './googleSheets';

const START_DATE = new Date('2026-10-01T00:00:00.000Z');

function makeHash(date: string, amount: number, accountNumber = '1234') {
  return `${date}_${amount}_Coffee_undefined_max_${accountNumber}`;
}

function makeTransaction(hash: string, status: 'pending' | 'completed'): EnrichedTransaction {
  return { hash, status, processedDate: '2026-11-10T00:00:00.000Z' } as EnrichedTransaction;
}

describe('getPendingRowChanges', () => {
  const pendingHash = makeHash('2026-10-20T21:00:00.000Z', -50);

  test('marks a pending row as completed when the same transaction is now completed', () => {
    const completed = makeTransaction(pendingHash, 'completed');
    const result = getPendingRowChanges(
      [{ rowIndex: 5, hash: pendingHash, status: 'pending' }],
      [completed],
      START_DATE,
    );

    expect(result.rowsToComplete).toEqual([{ rowIndex: 5, transaction: completed }]);
    expect(result.rowIndexesToDelete).toEqual([]);
  });

  test('leaves a pending row alone while the transaction is still pending', () => {
    const result = getPendingRowChanges(
      [{ rowIndex: 5, hash: pendingHash, status: 'pending' }],
      [makeTransaction(pendingHash, 'pending')],
      START_DATE,
    );

    expect(result).toEqual({ rowsToComplete: [], rowIndexesToDelete: [] });
  });

  test('deletes a pending row that no longer appears in a scrape of the same account', () => {
    const capturedWithNewAmount = makeTransaction(makeHash('2026-10-20T21:00:00.000Z', -52.3), 'completed');
    const result = getPendingRowChanges(
      [{ rowIndex: 5, hash: pendingHash, status: 'pending' }],
      [capturedWithNewAmount],
      START_DATE,
    );

    expect(result.rowIndexesToDelete).toEqual([5]);
  });

  test('never touches completed rows, rows of accounts that were not scraped, or rows before the start date', () => {
    const otherAccountHash = makeHash('2026-10-20T21:00:00.000Z', -50, '9999');
    const beforeStartHash = makeHash('2026-09-15T21:00:00.000Z', -50);
    const onStartBoundaryHash = makeHash('2026-10-01T10:00:00.000Z', -50);
    const result = getPendingRowChanges(
      [
        { rowIndex: 0, hash: 'hash - לא לגעת', status: 'סטטוס' },
        { rowIndex: 1, hash: pendingHash, status: 'completed' },
        { rowIndex: 2, hash: otherAccountHash, status: 'pending' },
        { rowIndex: 3, hash: beforeStartHash, status: 'pending' },
        { rowIndex: 4, hash: onStartBoundaryHash, status: 'pending' },
      ],
      [makeTransaction(makeHash('2026-10-22T21:00:00.000Z', -10), 'completed')],
      START_DATE,
    );

    expect(result).toEqual({ rowsToComplete: [], rowIndexesToDelete: [] });
  });
});
