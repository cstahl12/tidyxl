context("xlsx_cells()")

test_that("xlsx_cells() warns of missing sheets", {
  expect_warning(expect_error(xlsx_cells("./examples.xlsx", c(NA, NA)),"All elements of argument 'sheets' were discarded."),  "Argument 'sheets' included NAs, which were discarded.")
  expect_error(xlsx_cells("./examples.xlsx", "foo"), "Sheet\\(s\\) not found: \"foo\"")
  expect_error(xlsx_cells("./examples.xlsx", c("foo", "bar")), "Sheet\\(s\\) not found: \"foo\", \"bar\"")
  expect_error(xlsx_cells("./examples.xlsx", 5), "Only 3 sheet\\(s\\) found.")
  expect_error(xlsx_cells("./examples.xlsx", TRUE), "Argument `sheet` must be either an integer or a string.")
})

test_that("xlsx_cells() finds named sheets", {
  expect_error(xlsx_cells("./examples.xlsx", "Sheet1"), NA)
})

test_that("xlsx_cells() gracefully fails on missing files", {
  expect_error(xlsx_cells("foo.xlsx"), "'foo\\.xlsx' does not exist in current working directory \\('.*'\\).")
})

test_that("xlsx_cells() allows user interruptions", {
  # This is just for code coverage of the branch that checks for interruptions.
  # It doesn't attempt to interrupt.
  expect_error(xlsx_cells("./thousand.xlsx"), NA)
})

test_that("the highest cell address is parsed by xlsx_cells()", {
  expect_error(xlsx_cells("./xfd1048576.xlsx"), NA)
})
