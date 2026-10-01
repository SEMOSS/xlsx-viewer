import { TriangleAlert } from 'lucide-react';
import { useEffect, useMemo, useRef, useState, type KeyboardEvent } from 'react';
import type { RangeAddress, TextColumnFormat, TextToColumnsOptions } from 'xlsx-viewer-core';
import {
	guessFixedWidthBreaks,
	parseTextToColumns,
	textToColumnsOverwrites,
} from 'xlsx-viewer-core';

import { formatNumber, useLocale, useT } from '../../../i18n';
import { textToColumnsCmd } from '../../commands/data-tools-commands';
import { useViewerStore } from '../../state/use-store';
import { Dialog, DialogButton } from '../ui/Dialog';
import { DialogMessage } from './dialog-ui';
import {
	columnCountOf,
	DEFAULT_DELIMITERS,
	type DelimiterState,
	formatDestination,
	guessSplitMode,
	parseDestination,
	splitOptionsOf,
	Step1,
	Step3,
	StepDelimited,
	StepFixedWidth,
	textToColumnsSetup,
} from './text-to-columns';
import type { DialogProps } from './types';

type Step = 1 | 2 | 3;

/** Elements that handle Enter themselves. */
function handlesEnter(target: HTMLElement): boolean {
	return (
		Boolean(target.closest('[data-enter="self"]')) ||
		['BUTTON', 'TEXTAREA', 'A'].includes(target.tagName)
	);
}

function Alert({ message }: { message: string }) {
	return (
		<div className='flex items-start gap-3'>
			<TriangleAlert size={28} className='shrink-0 text-amber-500' aria-hidden />
			<p className='pt-1'>{message}</p>
		</div>
	);
}

/** Excel's message box shown instead of the wizard (several columns selected, no data). */
function ExcelMessage({ message, onClose }: { message: string; onClose: () => void }) {
	const t = useT();
	const store = useViewerStore();
	useEffect(() => {
		store.setUi({ message });
	}, [store, message]);
	return (
		<Dialog
			title={t('dialog.textToColumns.messageTitle')}
			onClose={onClose}
			width={440}
			initialFocusSelector='button[data-primary]'
			footer={
				<DialogButton primary onClick={onClose}>
					{t('dialog.ok')}
				</DialogButton>
			}
		>
			<div role='alert'>
				<Alert message={message} />
			</div>
		</Dialog>
	);
}

function Wizard({
	source,
	lines,
	onClose,
}: {
	source: RangeAddress;
	lines: string[];
	onClose: () => void;
}) {
	const t = useT();
	const locale = useLocale();
	const store = useViewerStore();
	const [guessed] = useState(() => guessSplitMode(lines));
	const [step, setStep] = useState<Step>(1);
	const [mode, setMode] = useState(guessed);
	const [delimiters, setDelimiters] = useState<DelimiterState>(DEFAULT_DELIMITERS);
	const [breaks, setBreaks] = useState(() => guessFixedWidthBreaks(lines));
	const [formats, setFormats] = useState<TextColumnFormat[]>([]);
	const [column, setColumn] = useState(0);
	const [destination, setDestination] = useState(() => formatDestination(source.top, source.left));
	const [error, setError] = useState<string | null>(null);
	const [confirm, setConfirm] = useState<TextToColumnsOptions | null>(null);
	const body = useRef<HTMLDivElement>(null);

	const split = useMemo(() => splitOptionsOf(mode, delimiters, breaks), [mode, delimiters, breaks]);
	const rows = useMemo(() => parseTextToColumns(lines, split), [lines, split]);
	const columns = columnCountOf(rows);
	const columnFormats = Array.from({ length: columns }, (_, i) => formats[i] ?? 'general');
	const selected = Math.min(column, columns - 1);

	useEffect(() => {
		const target = confirm
			? 'button[data-primary]'
			: 'input[type="radio"]:checked, input:not([type="radio"]), select';
		body.current?.querySelector<HTMLElement>(target)?.focus();
	}, [step, confirm]);

	const go = (to: Step) => {
		setError(null);
		setStep(to);
	};
	const apply = (options: TextToColumnsOptions) => {
		setConfirm(null);
		const message = textToColumnsCmd(store, source, options);
		if (message) {
			setError(message);
			return;
		}
		onClose();
	};
	const finish = () => {
		if (store.readOnly) {
			return;
		}
		const dest = parseDestination(destination, store.activeSheet.name);
		if (!dest) {
			setStep(3);
			setError(t('message.invalidReference'));
			return;
		}
		const options: TextToColumnsOptions = { ...split, destination: dest, columnFormats };
		if (textToColumnsOverwrites(store.workbook, store.state.activeSheet, source, options)) {
			setConfirm(options);
			return;
		}
		apply(options);
	};
	const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
		if (confirm && e.key === 'Escape') {
			// Escape answers the question (Cancel) instead of closing the wizard.
			e.stopPropagation();
			setConfirm(null);
			return;
		}
		if (e.key !== 'Enter' || e.shiftKey || e.altKey || e.ctrlKey || e.metaKey) {
			return;
		}
		if (e.nativeEvent.isComposing || handlesEnter(e.target as HTMLElement)) {
			return;
		}
		e.preventDefault();
		if (confirm) {
			apply(confirm);
		} else if (step < 3) {
			go((step + 1) as Step);
		} else {
			finish();
		}
	};

	return (
		<Dialog
			title={t('dialog.textToColumns.title', {
				step: formatNumber(step, locale),
				count: formatNumber(3, locale),
			})}
			onClose={onClose}
			width={640}
			footer={
				<>
					<DialogButton onClick={onClose} disabled={confirm !== null}>
						{t('dialog.cancel')}
					</DialogButton>
					<DialogButton
						onClick={() => go((step - 1) as Step)}
						disabled={step === 1 || confirm !== null}
					>
						{t('dialog.textToColumns.back')}
					</DialogButton>
					<DialogButton
						primary={step < 3}
						onClick={() => go((step + 1) as Step)}
						disabled={step === 3 || confirm !== null}
					>
						{t('dialog.textToColumns.next')}
					</DialogButton>
					<DialogButton
						primary={step === 3}
						onClick={finish}
						disabled={store.readOnly || confirm !== null}
					>
						{t('dialog.textToColumns.finish')}
					</DialogButton>
				</>
			}
		>
			<div ref={body} className='flex flex-col gap-3' onKeyDown={onKeyDown}>
				{confirm ? (
					<div
						role='alertdialog'
						aria-label={t('dialog.textToColumns.messageTitle')}
						className='flex flex-col gap-3 rounded-md border border-border p-3'
					>
						<Alert message={t('dialog.textToColumns.replaceData')} />
						<div className='flex justify-end gap-2'>
							<DialogButton onClick={() => setConfirm(null)}>{t('dialog.cancel')}</DialogButton>
							<DialogButton primary onClick={() => apply(confirm)}>
								{t('dialog.ok')}
							</DialogButton>
						</div>
					</div>
				) : (
					<>
						{step === 1 && <Step1 mode={mode} guessed={guessed} lines={lines} onMode={setMode} />}
						{step === 2 && mode === 'delimited' && (
							<StepDelimited
								value={delimiters}
								onChange={(patch) => setDelimiters((d) => ({ ...d, ...patch }))}
								rows={rows}
								columns={columns}
							/>
						)}
						{step === 2 && mode === 'fixed' && (
							<StepFixedWidth lines={lines} breaks={breaks} onChange={setBreaks} />
						)}
						{step === 3 && (
							<Step3
								formats={columnFormats}
								selected={selected}
								onSelect={setColumn}
								onFormat={(format) =>
									setFormats(columnFormats.map((f, i) => (i === selected ? format : f)))
								}
								destination={destination}
								onDestination={(text) => {
									setDestination(text);
									setError(null);
								}}
								rows={rows}
							/>
						)}
					</>
				)}
				{error && <DialogMessage>{error}</DialogMessage>}
			</div>
		</Dialog>
	);
}

/**
 * Data > Text to Columns: Excel's Convert Text to Columns Wizard for the
 * selected column (kind `textToColumns`).
 */
export function TextToColumnsDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const [setup] = useState(() => textToColumnsSetup(store));
	if (!setup.source) {
		const message = setup.error ?? t('dialog.textToColumns.noData');
		return <ExcelMessage message={message} onClose={onClose} />;
	}
	return <Wizard source={setup.source} lines={setup.lines} onClose={onClose} />;
}
