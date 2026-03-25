// quote-calculator.js

class QuoteCalculator {
    constructor(distance, weekdayRate, weekendRate, accommodationCost, optionalServices) {
        this.distance = distance; // in km
        this.weekdayRate = weekdayRate; // cost per km on weekdays
        this.weekendRate = weekendRate; // cost per km on weekends
        this.accommodationCost = accommodationCost; // fixed cost for accommodation
        this.optionalServices = optionalServices; // array of service costs
        this.taxRate = 0.1; // 10% tax
    }

    calculateTransportationCost(isWeekend) {
        const rate = isWeekend ? this.weekendRate : this.weekdayRate;
        return this.distance * rate;
    }

    calculateOptionalServicesCost() {
        return this.optionalServices.reduce((total, serviceCost) => total + serviceCost, 0);
    }

    calculateTotalCost(isWeekend) {
        const transportationCost = this.calculateTransportationCost(isWeekend);
        const optionalServicesCost = this.calculateOptionalServicesCost();
        const subtotal = transportationCost + this.accommodationCost + optionalServicesCost;
        const tax = subtotal * this.taxRate;
        return subtotal + tax;
    }

    getQuote(isWeekend) {
        const totalCost = this.calculateTotalCost(isWeekend);
        return {
            totalCost: totalCost.toFixed(2),
            breakdown: {
                transportationCost: this.calculateTransportationCost(isWeekend).toFixed(2),
                accommodationCost: this.accommodationCost.toFixed(2),
                optionalServicesCost: this.calculateOptionalServicesCost().toFixed(2),
                tax: (totalCost * this.taxRate).toFixed(2)
            }
        };
    }
}

// Usage example:
// const quote = new QuoteCalculator(100, 2.5, 3.0, 150, [50, 20]);
// const result = quote.getQuote(false); // false for weekday
// console.log(result);