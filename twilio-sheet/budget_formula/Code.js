/**
 * Calculates the total annual amount of an expense
 * by multplying its amount by it's frequency.
 *
 * @param {number} amount - dollar cost of an expense.
 * @param {number} frequency - annual, monthly, or weekly.
 *
 * @returns {number} - annual cost
 */
function ANNUAL_COST(amount, frequency) {
    return amount * timesPerYear(frequency);
  }

  /**
   * @param {string} frequency - frequency as a string.
   *
   * @return {number} - number of times per year.
   */
  function timesPerYear(frequency) {
    switch (frequency) {
      case '1 time':
        return 1;
      case 'Weekly':
        return 52;
      case '2 weeks':
        return 52 / 2;
      case 'Monthly':
        return 12;
      case '2 months':
        return 6;
      case 'Semester (6 months)':
        return 2;
      case 'Yearly':
        return 1;
    }
  }
