Excel Libraries and Tools
=========================

![Build Status](https://github.com/peterbecker/xls-utils/workflows/Build/badge.svg)

This is a collection of libraries and tools to work with Excel
files from Kotlin, building on the [Apache POI library](https://poi.apache.org/).

Contents:

* `kotlin-xls` - a wrapper around POI to make handling Excel files easier, also includes code to compare
   Excel workbooks, aimed primarily at unit testing
* `db2xls` - a command line utility to run an SQL query via JDBC and write the results into an Excel file
* `db2xls-bundle` - a version of `db2xls` bundled into a single JAR with all dependencies and some common
  JDBC drivers


Quick Start
===========

Finding or building the tool
----------------------------

There is currently no fully published version of this tool, but you can find the result of the most recent builds
in the "Artifacts" section of the builds in the "Action" tab on Github. The documentation on running the example
assumes you downloaded or checkout out a copy of the source code either way.

To build the command line tool, you need to build this project first using:

```shell script
mvn package
```

This requires a [Java Development Kit (JDK) 11 or greater](https://adoptopenjdk.net/), and 
[Maven 3.5](https://maven.apache.org/) or greater.

Running the Example
-------------------

Once the build has finished, go into `db2xls-bundle/sample` -- this contains a fully functional example with a
script `run.sh` to execute it (it's a single command, if you can't use the script run it directly).

If you downloaded the tool, copy the file into this folder and run:

```shell script
java -jar db2xls-bundle.jar countries.conf
```

This will:
* load the configuration from the `countries.conf` file
* load the template `countries.template.xlsx` as per configuration file
* open the database specified in the configuration -- this is an embedded database, which will be initialized using the
  `init.sql` script, which contains data about countries
* run queries as specified in the top left cells of the `Query_...` sheets
* fill in the tables with matching names (e.g. the query from `Query_Population` will fill the table named `Population`)
* write the result to `Countries.xlsx` as per configuration file

Running Other Queries
---------------------

The bundled version of this tool includes a number of database driver, currently:

* H2 (the embedded engine used in the example)
* PostgreSQL
* MySQL/MariaDB
* MS SQL Server

This means that the tool can connect to each of these by configuring a matching JDBC URL in the configuration file,
along with the username and password for the connection.

The only other two options are the input template and the name for the output file.

To set up your own template just take any Excel file and add new sheets named "Query_" followed by a name to identify
the query. Note that the prefix and the name are case sensitive, i.e. you need the `Q` capitalized and the name for
your query will have to match in capitalization as well. You can create any number of these query sheets, the tool
does not limit the number, although a very large number may create files that are too much for Excel to handle.

For each query create an Excel table ("Format as Table" in the Home ribbon). Then rename the table to match the query
name, e.g. if you created a sheet "Query_Recent_Sales", then the table should be named "Recent_Sales". Note that there
are limitations on table names, see [the documentation on renaming tables in Excel](https://support.microsoft.com/en-us/office/rename-an-excel-table-fbf49a4f-82a3-43eb-8ba2-44d21233b114). 

The table needs
to have the right width for the query result, so if the query results in 9 columns, then the table should be 9 columns
wide -- the tool will add rows as needed, but it will not change the width of the table. No rows will be 
deleted, so the table in the template should not contain data. Ideally it is only a header and a single empty data row.

You can also use named ranges
as targets, but tables produce nicer looking results and allow for easy sorting and filtering afterwards.

The target tables/ranges can be anywhere in the workbook, but the sheets named "Query_*" will be deleted after
the query was executed, so you need at least one more sheet for the output.

Once that is done, name the input template in the configuration file, give the output file a name, and run the tool
via this command:

```shell script
java -jar db2xls-bundle*.jar MY_CONFIG.conf
```

If you have more than one version of the tool, expand the `*` to the one you want to use.

Deployment
==========

To run the tool on another machine a Java 11+ runtime is required plus the following files:
* the bundle JAR file
* the configuration file
* the template file


Using other Database Systems
============================

If you need to query a database engine not supported out of the box, you also need the matching JDBC driver for this. 
Most RDBMS offer such a driver. You then need to run the following command instead:

```shell script
java -cp db2xls-bundle*.jar:JBDC_DRIVER_FILE.jar de.peterbecker.xls.Db2XlsKt MY_CONFIG.conf
```

On Windows the separator has to be a semicolon instead of the colon. Alternatively you can edit `db2xls-bundle/pom.xml`
to tailor the JDBC drivers in the file -- you can remove any you don't need this way as well. Note that the ones
specified in this file will pick up a version from the `pom.xml` in the top folder, you will have to specify the
version when adding custom ones, but this can be done directly in `db2xls-bundle/pom.xml`.

Programmatic Use
================

The logic behind the command line tool is all behind a single method in the `db2xls` module:

```
de.peterbecker.xls.Db2Xls::runReports(..)
```

This operates on POI `Workbook` objects, it takes the template and a JDBC connection and returns the modified workbook.
Note that at the time of writing it will operate on the input, so the template will have to be copied or loaded before
each use.

The workbook objects can be written to any output stream, which can be used to upload files somewhere or to use them
as a response to a HTTP request.

The `kotlin-xls` library contains lower-level functions that are not relating to database interactions. This includes
the ability to write in-memory data into the named tables or ranges.
