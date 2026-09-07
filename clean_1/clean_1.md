# Power Query Data Cleaning Project

## Overview

The main data-cleaning activities carried out in this project were:

- Cleaning and standardizing inconsistent `Order Date` values
- Merging related datasets using a common field
- Preparing the final dataset for analysis

A major challenge during the cleaning process was that the `Order Date` column contained dates in different formats. Rather than simply applying a standard date type, a custom Power Query M transformation was used to identify the structure of each date and convert it where possible.

## 1. Cleaning the Order Date Column

### Problem Identified

The `Order Date` column contained date values that were not consistently formatted.

Dates could appear in different arrangements, for example:

- `2024-03-15`
- `15/03/2024`
- `03/15/2024`
- `15-03-2024`
- `03-15-2024`

This created a problem because Power Query could potentially interpret the same date differently depending on the format.

For example, `03/04/2024` could represent either:

- 3 April 2024 (DD/MM/YYYY)
- 4 March 2024 (MM/DD/YYYY)

Therefore, the dates needed to be examined and interpreted before being converted into a proper date value.

## 2. Custom Power Query M Date-Cleaning Logic

The following custom M code was used to clean the `Order Date` column:

```m
let
    x = if [Order Date] = null then null else Text.Trim(Text.From([Order Date])),
    parts = 
        if x = null then {}
        else if Text.Contains(x, "/") then Text.Split(x, "/")
        else Text.Split(x, "-"),

    result =
        if List.Count(parts) <> 3 then
            x
        else
            let
                a = Number.FromText(parts{0}),
                b = Number.FromText(parts{1}),
                c = Number.FromText(parts{2})
            in
                if Text.Length(parts{0}) = 4 then
                    #date(a,b,c)
                else if a > 12 then
                    #date(c,b,a)
                else if b > 12 then
                    #date(c,a,b)
                else
                    x
in
    result
```

This transformation was important because it did not assume that every date followed the same format. Instead, it used the values themselves to determine how the date should be interpreted.

## 3. Explanation of the Date-Cleaning Code

### Step 1: Handling Null Values

The first part of the code checks whether `Order Date` is null:

```m
x = if [Order Date] = null
    then null
    else Text.Trim(Text.From([Order Date]))
```

- If the value is missing, it remains `null`.
- Otherwise, the value is converted to text, and `Text.Trim()` removes unnecessary spaces.

This helps prevent formatting problems caused by values such as `" 15/03/2024 "`.

### Step 2: Splitting the Date

The code checks whether the date uses `/` or `-` as a separator:

```m
parts = 
    if x = null then {}
    else if Text.Contains(x, "/") then Text.Split(x, "/")
    else Text.Split(x, "-")
```

For example, `15/03/2024` becomes:

```
{"15", "03", "2024"}
```

Similarly, `2024-03-15` becomes:

```
{"2024", "03", "15"}
```

This allows each component of the date to be examined individually.

## 4. Validating the Number of Date Components

The code checks whether the date contains exactly three components:

```m
if List.Count(parts) <> 3 then
    x
```

A normal date should contain three parts, in one of these orders:

- Day / Month / Year
- Month / Day / Year
- Year / Month / Day

If the value does not contain three parts, the original value is returned rather than forcing an incorrect conversion.

## 5. Converting Date Components to Numbers

The three parts of the date are assigned to variables:

```m
a = Number.FromText(parts{0}),
b = Number.FromText(parts{1}),
c = Number.FromText(parts{2})
```

The variables represent the three components of the date. For example, `15/03/2024` becomes approximately:

- `a = 15`
- `b = 3`
- `c = 2024`

This allows the code to determine which component represents the day, month, and year.

## 6. Detecting the YYYY-MM-DD Format

The first major condition is:

```m
if Text.Length(parts{0}) = 4 then
    #date(a,b,c)
```

If the first component contains four digits, it is assumed to be the year.

For example, `2024-03-15` is interpreted as:

- Year = 2024
- Month = 03
- Day = 15

The `#date()` function then creates a proper Power Query date: `#date(2024, 3, 15)`.

## 7. Detecting DD/MM/YYYY

The next condition is:

```m
else if a > 12 then
    #date(c,b,a)
```

If the first value is greater than 12, it cannot represent a month.

For example, in `15/03/2024`, since `15 > 12`, the code recognizes 15 as the day. Therefore, the date is interpreted as:

- Day = 15
- Month = 03
- Year = 2024

and converted using `#date(c,b,a)`.

**Result:** `2024-03-15`

## 8. Detecting MM/DD/YYYY

The next condition is:

```m
else if b > 12 then
    #date(c,a,b)
```

If the second value is greater than 12, it cannot represent a month.

For example, in `03/15/2024`, since `15 > 12`, the second value must be the day. Therefore:

- Month = 03
- Day = 15
- Year = 2024

The code converts it using `#date(c,a,b)`.

**Result:** `2024-03-15`

## 9. Handling Ambiguous Dates

The final condition is:

```m
else
    x
```

This is particularly important. Consider `03/04/2024` — both `03` and `04` are less than or equal to 12. Therefore, the code cannot reliably determine whether the date means:

- 03 April 2024, or
- 04 March 2024

Instead of making an assumption that could introduce incorrect data, the code leaves the original value unchanged. This is a safer approach when the original dataset does not provide enough information to distinguish between DD/MM/YYYY and MM/DD/YYYY.

## 10. Date Formatting Outcome

The custom transformation converts dates that can be reliably identified into proper date values.

| Original Value | Interpretation | Result |
|---|---|---|
| `2024-03-15` | YYYY-MM-DD | `2024-03-15` |
| `15/03/2024` | DD/MM/YYYY | `2024-03-15` |
| `03/15/2024` | MM/DD/YYYY | `2024-03-15` |
| `15-03-2024` | DD-MM-YYYY | `2024-03-15` |
| `03-15-2024` | MM-DD-YYYY | `2024-03-15` |
| `03/04/2024` | Ambiguous | Original value retained |
| `null` | Missing | `null` |

The important point is that the transformation does not blindly convert every value. It applies rules based on the structure and numerical values of the date.

## 11. Merging Datasets

After cleaning the date information, the related datasets were merged in Power Query.

The merge was performed using a common field between the datasets, such as an ID or another matching key.

A typical Power Query merge operation can be represented using:

```m
Table.NestedJoin(
    MainTable,
    {"ID"},
    SecondTable,
    {"ID"},
    "SecondTable",
    JoinKind.LeftOuter
)
```

A `LeftOuter` join retains all records from the main table while bringing in matching information from the second table.

After the merge, the required columns from the second dataset can be expanded using:

```m
Table.ExpandTableColumn(
    MergedTable,
    "SecondTable",
    {"Column1", "Column2"},
    {"Column1", "Column2"}
)
```

This made it possible to combine information from the separate datasets into one analysis-ready table.

## 12. Why These Cleaning Steps Were Important

The date-cleaning process was important because inconsistent dates can lead to incorrect:

- Sorting
- Filtering
- Grouping
- Time-based calculations
- Monthly or yearly analysis
- Trend analysis

The merging process was also important because relevant information was stored across multiple datasets. Combining the datasets allowed the information to be analyzed together.

## 13. Final Data-Cleaning Workflow

The overall Power Query workflow was:

```
Raw Data
   ↓
Load Data into Power Query
   ↓
Inspect Order Date
   ↓
Clean and Interpret Inconsistent Dates
   ↓
Convert Valid Values to Date
   ↓
Merge Related Datasets
   ↓
Expand Required Columns
   ↓
Check the Result
   ↓
Load Cleaned Data for Analysis
```

## Conclusion

The most prominent data-cleaning work performed in this Power Query project involved handling inconsistent date formats and merging related datasets.

The custom M transformation was particularly important because the `Order Date` column contained dates represented using different separators and date arrangements. The code examined each value, identified recognizable date patterns, and converted them into valid date values where the format could be determined reliably.

The datasets were then merged using a common field to create a more complete dataset for analysis.

Overall, the Power Query process improved the consistency, reliability, and usability of the data while creating a repeatable transformation workflow.
