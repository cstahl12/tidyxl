context("xlsx_formats()")

test_that("all known formats don't break xlsx_formats()", {
  expect_error(xlsx_formats("./examples.xlsx"), NA)
})

