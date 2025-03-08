# Excel Lambdas

A collection of useful Excel lambdas for streamlining complex tasks

## ReplicateArray

The **ReplicateArray** lambda function replicates each element of an array. It handles both one-dimensional arrays (rows or columns) and two-dimensional arrays. For a 2D array, it copies each element in both dimensions. For a 1D array, it replicates elements along its single dimension.

### Code

```excel
=LAMBDA(array, copies,
  LET(
    rows, ROWS(array),
    cols, COLUMNS(array),
    result,
      IF(AND(rows>1, cols>1),
         /* Two-dimensional array: replicate in both rows and columns */
         LET(
           newRows, rows * copies,
           newCols, cols * copies,
           MAKEARRAY(newRows, newCols,
             LAMBDA(r, c,
               INDEX(array, INT((r-1)/copies)+1, INT((c-1)/copies)+1)
             )
           )
         ),
         /* One-dimensional array: replicate along its only dimension */
         IF(rows=1,
           LET(
             len, cols,
             indices, INT((SEQUENCE(1, len * copies)-1)/copies)+1,
             INDEX(array, 1, indices)
           ),
           LET(
             len, rows,
             indices, INT((SEQUENCE(len * copies)-1)/copies)+1,
             INDEX(array, indices, 1)
           )
         )
      ),
    result
  )
)
```
