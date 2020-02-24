// Row indices.
const CUSTOMER_ID = 0;
const DATE = 1;
const AMOUNT = 2;

const MILLISECONDS_IN_MONTH = 1000 * 60 * 60 * 24 * 30;

/**
 * Groups customers by their ID calculates
 * customer lifetime value for each.
 *
 * @param {Any[][]} rows - array of customer ID, date, amount.
 *
 * @returns {Any[][]} - table, each row containing:
 *    - average order size
 *    - avgOrderFrequency,
          avgCustomerValue,
          avgCustomerLifespan,
          customerLifetimeValue,
 *         map of id:orders
 */
function LIFETIME_VALUES(rows) {
  // Group all the orders per customer ID.
  // After this we get all the customers with a list of all
  // their associated orders.
  let ordersPerCustomer = groupByKey(rows
      // Keep only rows that don't have a blank customer ID.
      .filter(row => !!row[CUSTOMER_ID])
      // Create [key, value] pairs using the customer ID as key.
      .map(row => [row[CUSTOMER_ID], {
        date: row[DATE],
        amount: row[AMOUNT],
      }]));

  return ordersPerCustomer
      // Create objects from the [key, value] pairs.
      .map(keyValue => ({
        id: keyValue[0],
        orders: keyValue[1],
      }))
      // Calculate the values we need and return an array of
      // the columns to populate per element.
      .map(customer => {
        let totalRevenue = customer.orders
            // Get the `amount` field from each order.
            .map(order => order.amount)
            // Sum all the amounts.
            .reduce((result, amount) => result + amount);

        let datesRange = customer.orders
            // Get the `date` field from each order.
            .map(order => order.date)
            // Compares each date and assigns the minimum and maximum.
            // Since we want to return an object with {first: ..., last: ...}
            // instead of a date, we must set an initial value it explicitly.
            .reduce((result, date) => ({
              first: Math.min(result.first || date, date),
              last: Math.max(result.last || date, date),
            }), {first: null, last: null});

        // Dates are stored as UNIX timestamps in milliseconds.
        // https://en.wikipedia.org/wiki/Unix_time
        let totalMonths = (datesRange.last - datesRange.first) / MILLISECONDS_IN_MONTH;

        // Calculate the values we are interested in.
        let avgOrderSize = totalRevenue / customer.orders.length;
        let avgOrderFrequency = customer.orders.length / ordersPerCustomer.length;
        let avgCustomerValue = avgOrderSize * avgOrderFrequency;
        let avgCustomerLifespan = totalMonths;
        let customerLifetimeValue = avgCustomerValue * avgCustomerLifespan;

        // Returning an array makes the custom function to fill multiple cells
        // instead of a single cell.
        // In this case, we are returning one array for every order in
        // `ordersPerCustomer`, so the custom function fills a 2D matrix of cells.
        // User must create these headers manually in the spreadsheet in this order.
        return [
          customer.id,
          avgOrderSize,
          avgOrderFrequency,
          avgCustomerValue,
          avgCustomerLifespan,
          customerLifetimeValue,
        ]
      });
}

/**
 * Groups elements of an array by key.
 *
 * @param {Any[][]} keyValues - array of [key, value] pairs.
 *
 * @returns {Any[][]} - array of [key, values[]] pairs, where each
 *      key is unique, and is mapped to all its values.
 *
 */
function groupByKey(keyValues) {
  // Reduce the array of [key, value] pairs into a Map with unique
  // keys, and an array of all the values associated to that key.
  let keyValuesMap = keyValues.reduce((result, keyValue) => {
    let [key, value] = keyValue;
    if (!result.has(key))
      result.set(key, []);
    result.get(key).push(value);
    return result;
  }, new Map());

  // Map.entries() gives us an Iterator of [key, value] pairs.
  // We use the spread operator to create an array with its contents.
  // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Spread_syntax
  return [...keyValuesMap.entries()];
}
