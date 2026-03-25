// Quote calculation engine for quote-automation project

class QuoteCalculator {
    constructor(basePrice) {
        this.basePrice = basePrice;
    }

    calculateTax(taxRate) {
        return this.basePrice * (taxRate / 100);
    }

    calculateTotal(taxRate, discount) {
        const tax = this.calculateTax(taxRate);
        const total = this.basePrice + tax - discount;
        return total;
    }
}

// Example usage:
const calculator = new QuoteCalculator(1000);
const taxRate = 5; // 5%
const discount = 50;
const totalPrice = calculator.calculateTotal(taxRate, discount);
console.log(`Total Price: $${totalPrice}`);