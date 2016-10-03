## Weather webpart simple extended
This webpart demonstrates functionality built on top of the webpart provided by Waldek - https://github.com/SharePoint/sp-dev-fx-webparts/tree/master/samples/jquery-cdn
![alt tag](https://cloud.githubusercontent.com/assets/9694225/19053382/9e821f64-89b2-11e6-8054-d9c52518aa9c.gif)
This webpart provides current weather and forecast upto 5 days of:
- A pre-configured location from a list (Bangalore)
- A location specified in the webpart property (New York)
- Location picked form current user's user profile.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* commonjs components - this allows this package to be reused from other packages.
* dist/* - a single bundle containing the components used for uploading to a cdn pointing a registered Sharepoint webpart library to.
* example/* a test page that hosts all components in this package.

### Build options

gulp nuke - TODO
gulp test - TODO
gulp watch - TODO
gulp build - TODO
gulp deploy - TODO
