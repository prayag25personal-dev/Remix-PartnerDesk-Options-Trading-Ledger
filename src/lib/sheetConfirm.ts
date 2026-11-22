export interface SheetConfirmRequest {
  title: string;
  description: string;
  confirmLabel?: string;
  cancelLabel?: string;
  tone?: 'primary' | 'danger';
  details?: string[];
}

export const SHEET_ACTION_CANCELLED = 'SHEET_ACTION_CANCELLED';

type ConfirmHandler = (req: SheetConfirmRequest) => Promise<boolean>;

let registeredHandler: ConfirmHandler | null = null;

export function registerSheetConfirmHandler(handler: ConfirmHandler | null): void {
  registeredHandler = handler;
}

export async function confirmSheetAction(req: SheetConfirmRequest): Promise<void> {
  if (!registeredHandler) {
    throw new Error('Confirmation dialog is not ready. Please try again.');
  }
  const approved = await registeredHandler(req);
  if (!approved) {
    throw new Error(SHEET_ACTION_CANCELLED);
  }
}

export function isSheetActionCancelled(err: unknown): boolean {
  return err instanceof Error && err.message === SHEET_ACTION_CANCELLED;
}
