const DIGITS = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
const PLACES = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
const MILLION = 'ล้าน';

/** Thai words for 0 < n < 1,000,000. */
function belowMillion(n: number): string {
	const digits = String(n).split('').map(Number).reverse();
	let out = '';
	for (let i = digits.length - 1; i >= 0; i--) {
		const d = digits[i];
		if (d === 0) {
			continue;
		}
		if (i === 1) {
			out += d === 1 ? 'สิบ' : d === 2 ? 'ยี่สิบ' : `${DIGITS[d]}สิบ`;
		} else if (i === 0 && d === 1 && digits.length > 1 && digits.slice(1).some((x) => x !== 0)) {
			out += 'เอ็ด';
		} else {
			out += DIGITS[d] + PLACES[i];
		}
	}
	return out;
}

/** Thai words for a non-negative integer, grouping by millions (ล้าน). */
function integerWords(n: number): string {
	if (n === 0) {
		return '';
	}
	const groups: number[] = [];
	let rest = n;
	while (rest > 0) {
		groups.push(rest % 1_000_000);
		rest = Math.floor(rest / 1_000_000);
	}
	let out = '';
	for (let i = groups.length - 1; i >= 0; i--) {
		if (groups[i] > 0) {
			out += belowMillion(groups[i]);
		}
		if (i > 0) {
			out += MILLION;
		}
	}
	return out;
}

/**
 * BAHTTEXT: a number as Thai text in baht and satang, e.g. 1234.5 ->
 * หนึ่งพันสองร้อยสามสิบสี่บาทห้าสิบสตางค์. Whole amounts end in ถ้วน.
 */
export function bahtText(value: number): string {
	const negative = value < 0;
	const cents = Math.round(Math.abs(value) * 100);
	const baht = Math.floor(cents / 100);
	const satang = cents % 100;
	let out = '';
	if (baht === 0 && satang === 0) {
		out = 'ศูนย์บาทถ้วน';
	} else {
		if (baht > 0) {
			out += `${integerWords(baht)}บาท`;
		}
		out += satang > 0 ? `${integerWords(satang)}สตางค์` : 'ถ้วน';
	}
	return negative ? `ลบ${out}` : out;
}
