import { isDate, isString } from 'lodash';
import { escape } from '../utils';

export default function (value, cell) {
    if (isDate(value)) {
        const unixTimestamp = value.getTime();
        const officeTimestamp = (unixTimestamp / 86400000) + 25569;
        return `<c r="${cell}" t="n"><v>${officeTimestamp}</v></c>`;
    } else if (isString(value)) {
        return `<c r="${cell}" t="inlineStr"><is><t>${escape(value)}</t></is></c>`;
    } else if (value) {
        return `<c r="${cell}" t="n"><v>${value}</v></c>`;
    }
    return '';
}