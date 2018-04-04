## spfx-gantt

This is a web part built for use against default sharepoint task lists.

It uses dhtmlx gantt (https://dhtmlx.com/docs/products/dhtmlxGantt/) to render the tasks in a gantt view, and has full CRUD against the task list.

Some features:
* Drag tasks to update dates
* Set zoom to default Hours/Days/Months
* Uses the search api for the people picker
* New tasks and updated tasks sends email updates if there is an assigned user and if the checkbox is checked

Some work left:
* The dhtmlx gantt JS uses a global variable (gantt) - only one web part per page.
* Solution not 100% tested
* The control is not 100% React, I did not have enough time to do so :(

## DEMO
![Demo video](https://github.com/kmosti/spfx-gantt/blob/master/preview/SPFX_GANTT_v1.1.gif?raw=true)

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
