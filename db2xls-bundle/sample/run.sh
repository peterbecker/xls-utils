#!/usr/bin/env sh

## Note: this requires the build to be run first with `mvn package` at the folder above
java -jar ../target/db2xls-bundle-*.jar countries.conf
