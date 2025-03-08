# Excel Lambdas

A collection of useful Excel lambdas for streamlining complex tasks

## ReplicateArray

The **ReplicateArray** lambda function replicates each element of an array. It handles both one-dimensional arrays (rows or columns) and two-dimensional arrays. For a 2D array, it copies each element in both dimensions. For a 1D array, it replicates elements along its single dimension.

### Code

```excel
ReplicateArray = LAMBDA(array, copies,
  LET(
    rows, ROWS(array),
    cols, COLUMNS(array),
    IF(
      AND(rows > 1, cols > 1),
      /* Two-dimensional array: replicate in both rows and columns */
      LET(
        newRows, rows * copies,
        newCols, cols * copies,
        MAKEARRAY(
          newRows,
          newCols,
          LAMBDA(r, c,
            INDEX(
              array,
              INT((r - 1) / copies) + 1,
              INT((c - 1) / copies) + 1
            )
          )
        )
      ),
      /* One-dimensional array: replicate along its only dimension */
      IF(
        rows = 1,
        LET(
          len, cols,
          indices, INT((SEQUENCE(1, len * copies) - 1) / copies) + 1,
          INDEX(array, 1, indices)
        ),
        LET(
          len, rows,
          indices, INT((SEQUENCE(len * copies) - 1) / copies) + 1,
          INDEX(array, indices, 1)
        )
      )
    )
  )
);
```

## FillSpaces

The **FillSpaces** lambda function expands an array by inserting blank cells between the original elements. It takes an input array and a numeric argument that specifies the number of spaces to insert after each element. For each cell in the input array, it creates a block of size (spaces+1)×(spaces+1) where the original element is placed in the top‑left cell and the remaining cells are left blank. The function supports both one-dimensional arrays (rows or columns) and two-dimensional arrays.

### Code

```excel

FillSpaces = LAMBDA(array, spaces,
  LET(
    blockSize, spaces + 1,
    rows, ROWS(array),
    cols, COLUMNS(array),
    IF(
      AND(rows > 1, cols > 1),
      /* Two-dimensional array */
      LET(
        newRows, rows * blockSize,
        newCols, cols * blockSize,
        MAKEARRAY(
          newRows,
          newCols,
          LAMBDA(r, c,
            IF(
              AND(
                MOD(r - 1, blockSize) = 0,
                MOD(c - 1, blockSize) = 0
              ),
              INDEX(
                array,
                INT((r - 1) / blockSize) + 1,
                INT((c - 1) / blockSize) + 1
              ),
              ""
            )
          )
        )
      ),
      IF(
        rows = 1,
        /* Row vector */
        LET(
          newCols, cols * blockSize,
          MAKEARRAY(
            1,
            newCols,
            LAMBDA(r, c,
              IF(
                MOD(c - 1, blockSize) = 0,
                INDEX(array, 1, INT((c - 1) / blockSize) + 1),
                ""
              )
            )
          )
        ),
        /* Column vector */
        LET(
          newRows, rows * blockSize,
          MAKEARRAY(
            newRows,
            1,
            LAMBDA(r, c,
              IF(
                MOD(r - 1, blockSize) = 0,
                INDEX(array, INT((r - 1) / blockSize) + 1, 1),
                ""
              )
            )
          )
        )
      )
    )
  )
);
```