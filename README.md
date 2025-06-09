# Excel Formula Notes

Here are the useful formula that I have come acrossed:

### New Line Character

```swift
=CHAR(10)
```

## LET Keyword to Reuse Result

The `LET` function (available in Microsoft 365) allows you to assign names to calculations or values, making formulas more readable and efficient by reusing intermediate results. It reduces redundant calculations and clarifies complex formulas.

### Example

```swift
=LET(sum_result, SUM(A1:A10), sum_result*2 + sum_result/3)
```

In this example:

- `sum_result` is assigned the value of `SUM(A1:A10)`.
- The final result is `sum_result*2 + sum_result/3`, reusing `sum_result` to avoid recalculating the sum.

## TEXTJOIN for Joining Range with Delimiter

The `TEXTJOIN` function (available in Excel 2019 and Microsoft 365) combines text from a range or array, using a specified delimiter between each item. It’s useful for creating concatenated strings, such as comma-separated lists, and can ignore empty cells.

### Example

```swift
=TEXTJOIN(", ", TRUE, B1:B5)
```

In this example:

- `", "` is the delimiter (a comma followed by a space).
- `TRUE` ignores empty cells in the range `B1:B5`.
- If `B1:B5` contains `Apple`, `Banana`, `Cherry`, an empty cell, and `Date`, the result is `Apple, Banana, Cherry, Date`.

## UNIQUE to Remove Duplicate Items from a Range

The `UNIQUE` function (available in Microsoft 365) extracts unique values from a range or array, automatically removing duplicates. It’s ideal for data cleaning or creating lists of distinct items, and it supports dynamic updates as data changes.

### Example

```swift
=UNIQUE(C1:C10)
```

In this example:

- `C1:C10` might contain a list with duplicates, such as `Red`, `Blue`, `Red`, `Green`, `Blue`.
- `UNIQUE(C1:C10)` returns a dynamic array: `Red`, `Blue`, `Green`.
- The result spills into adjacent cells if entered in a single cell.

## & Operator to Join Values Row by Row

The `&` operator concatenates (joins) text or values in Excel, combining them into a single string. It’s commonly used to merge cell contents or add static text, processing data row by row in a formula.

### Example

```swift
=A1 & " - " & B1
```

In this example:

- If `A1` contains `John` and `B1` contains `Doe`, the formula combines them with a hyphen and space.
- The result is `John - Doe`.
- This can be copied down a column to join values for each row.

## NETWORKDAYS to Count Working Days Between Dates

The `NETWORKDAYS` function calculates the number of working days (excluding weekends and optionally specified holidays) between two dates. This is useful for project planning, payroll, or any scenario where you need to count business days.

### Example

```excel
=NETWORKDAYS(A1, B1, C1:C5)
```

In this example:

- `A1` is the start date.
- `B1` is the end date.
- `C1:C5` is an optional range containing holiday dates to exclude.
- The result is the count of working days between `A1` and `B1`, excluding weekends and any holidays listed in `C1:C5`.

## SUMPRODUCT for Conditional Counting or Summing

The `SUMPRODUCT` function multiplies corresponding components in given arrays and returns the sum of those products. It's often used for conditional counting or summing without needing array formulas.

### Example

```excel
=SUMPRODUCT((INT(Holidays[Date])>=A2)*(INT(Holidays[Date])<=EOMONTH(A2,0))*(WEEKDAY(INT(Holidays[Date]))=7))
```

In this example:

- `Holidays[Date]` refers to a column of dates in a table named `Holidays`.
- `A2` is the start date (typically the first day of a month).
- `EOMONTH(A2,0)` returns the last day of the month for the date in `A2`.
- `WEEKDAY(...)=7` checks if the date is a Saturday (if your system's week starts on Sunday).
- The formula counts how many Saturdays in the `Holidays[Date]` column fall within the month specified by `A2`.
