// test/test-quotes.js

const { calculateQuote } = require('../path/to/quote-module');

describe('Quote Calculations', () => {

    test('Calculate basic quote', () => {
        const result = calculateQuote(100);
        expect(result).toBe(100);
    });

    test('Calculate quote with discount', () => {
        const result = calculateQuote(100, { discount: 10 });
        expect(result).toBe(90);
    });

    test('Calculate quote with tax', () => {
        const result = calculateQuote(100, { tax: 0.2 });
        expect(result).toBe(120);
    });

    test('Calculate quote with discount and tax', () => {
        const result = calculateQuote(100, { discount: 10, tax: 0.2 });
        expect(result).toBe(108);
    });

});