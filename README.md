Excel Libraries and Tools
=========================

[![Build Status](https://travis-ci.org/peterbecker/xls-utils.svg?branch=master)](https://travis-ci.org/peterbecker/xls-utils)

This is (or rather 'hopefully will be') a collection of libraries and tools to work with Excel
files from Kotlin, building on the Apache POI library.

Contents:

* `kotlin-xls` - a wrapper around POI to make handling Excel files easier, also includes code to compare
   Excel workbooks, aimed primarily at unit testing
* `db2xls` - a command line utility to run an SQL query via JDBC and write the results into an Excel file
* `db2xls-bundle` - a version of `db2xls` bundled into a single JAR with all dependencies and some common
  JDBC drivers
