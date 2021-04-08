# Docxgen
Docxgen is a simple program that dynamically replaces certain markers in a template Word file with automatically generated or user-selected input.

## Simple Examples
```This is some text written by {name}```

This will prompt the user for the mandatory input `name`, and replace `{name}` with it.


```This is some text {*optional}.```

This will prompt the user for the optional input `\*optional`, replacing "{*optional}` with it. Note that in this context, there is no technical difference between a mandatory input and an optional one; the * simply denotes to the user that it is optional.


```This is some text{*}, here's even more: {*moreText}{/}.```

This will prompt the user for the optional input `\*moreText`, and replace `{\*moreText}` with it. If left blank, the entire block (denoted by `{*}` and `{/}`)is deleted, and the output reads "This is some text."
If multiple optional variables are included in the block, only one must be filled in for the block to be included. If they are all empty, however, the block is deleted.


## Replacement
Docxgen has `Replacement` abstract object that can be extended to provide additional, automatic functionality. The most simple is the `DateTimeReplacement` type:

```This document was written on {DATE} at {TIME}.```

Without asking for user input, the program will replace {DATE} and {TIME} with the date and time, respectively.


```Valid as of: {DATETIME}```

Without asking for user input, the program will replace {DATETIME} with the date and time.
