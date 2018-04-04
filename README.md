## spfx-gantt

This is a web part built for use against default sharepoint task lists.

It uses dhtmlx gantt (https://dhtmlx.com/docs/products/dhtmlxGantt/) to render the tasks in a gantt view, and has full CRUD against the task list.



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
