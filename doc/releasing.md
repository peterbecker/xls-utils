# Release Process

* Create a new release on GitHub at https://github.com/peterbecker/xls-utils/releases
  * "Draft new Release"
  * use version number for tag and release title
  * add description
  * publish -> should trigger a release build visible at https://github.com/peterbecker/xls-utils/actions
* Go to https://oss.sonatype.org/ and log in
* Once build is finished, the result should show in the "Staging Repositories" section (left hand menu)
* "Close" the release there (sic!) via the top menu
* wait for the closing to finish - you can see progress in the "Actions" tab. If something goes wrong, errors will show
  there, too.
* once the release is closed, "Release" it from the top menu
  * the release should then show as a released artifact in the Nexus UI on the site (e.g. when searching for 'kotlin-xls')
* wait until it appears at https://mvnrepository.com/artifact/com.github.peterbecker/kotlin-xls - this can take rather 
  long 