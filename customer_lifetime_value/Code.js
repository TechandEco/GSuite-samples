// Row indices.
const CUSTOMER_ID = 0;
const DATE = 1;
const AMOUNT = 2;

const MILLISECONDS_IN_MONTH = 1000 * 60 * 60 * 24 * 30;

function LIFETIME_VALUE(customerId, totalCustomers, customerDateAmountRows) {
  let orders = customerDateAmountRows
      // Filter and keep only the rows that match the customerId.
      .filter(row => row[0] == customerId)
      // Create an object from the row values.
      .map(row => ({
        date: row[1],
        amount: row[2],
      }));

  let totalRevenue = orders
      // Get the `amount` field from each order.
      .map(order => order.amount)
      // Sum all the amounts.
      .reduce((result, amount) => result + amount);

  let datesRange = orders
      // Get the `date` field from each order.
      .map(order => order.date)
      // Compares each date and assigns the minimum and maximum.
      // Since we want to return an object with {min: ..., max: ...}
      // instead of a date, we must set an initial value it explicitly.
      .reduce((result, date) => ({
        min: Math.min(result.min || date, date),
        max: Math.max(result.max || date, date),
      }), {min: null, max: null});

  // Dates are stored as UNIX timestamps in milliseconds.
  // https://en.wikipedia.org/wiki/Unix_time
  let totalMonths = (datesRange.max - datesRange.min) / MILLISECONDS_IN_MONTH;

  let avgOrderSize = totalRevenue / orders.length;
  let avgOrderFrequency = orders.length / totalCustomers;
  let avgCustomerValue = avgOrderSize * avgOrderFrequency;
  let avgCustomerLifespan = totalMonths;
  let customerLifetimeValue = avgCustomerValue * avgCustomerLifespan;

  return customerLifetimeValue;
}
