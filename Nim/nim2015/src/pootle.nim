import tables

var myTable = { "^": @[0, 1], "v": @[0, -1], "<": @[-1, 0], ">": @[1, 0] }.toTable

var myArray = myTable["^"]

