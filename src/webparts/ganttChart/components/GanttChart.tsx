/*global gantt*/
import * as React from 'react';
import styles from './GanttChart.module.scss';
import {
  IGanttChartProps,
  IGanttChartState,
  IGanttDataObject,
  IGanttData,
  IGanttLink,
  IGanttChartItemProp
 } from './IGanttChartProps';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';
import { Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Dialog,
  DialogType,
  PrimaryButton,
  IButtonProps,
  Label
} from 'office-ui-fabric-react';
import pnp, { ItemAddResult, SearchQueryBuilder, SearchResults, EmailProperties } from "sp-pnp-js";
import * as moment from 'moment';
const wysiwyg = require("wysiwyg");
require('dhtmlx');
require("gantt");
require("../../../../node_modules/dhtmlx-gantt/codebase/dhtmlxgantt.css");
require("./custom.css");
//require("./dhtmlx.css");

declare var gantt: any;
declare var dhtmlXComboFromSelect: any;
declare var dhtmlXCombo: any;

interface ganttOptions {
  options: peopleSearchResults[];
}

interface peopleSearchResults {
  value: string;
  text: string;
}

export default class GanttChart extends React.Component<IGanttChartProps, IGanttChartState> {

  constructor(props: IGanttChartProps, state: IGanttChartState) {
    super(props);
    this.state = {
      loading: true,
      error: "",
      results: null,
      showError: false,
      zoom: this.props.zoom,
      height: 200
    };
    this.handleZoomChange = this.handleZoomChange.bind(this);
  }

  private rowHeight: number = 35;

  public componentDidMount(): void {
    if (this.props.listTitle) {
      this._processInformation();
    }
  }
    /**
     * Called after a properties or state update
     * @param prevProps
     * @param prevState
     */
  public componentDidUpdate(prevProps:IGanttChartProps, prevState: IGanttChartState): void {
    if (prevState.zoom !== this.state.zoom) {
      //this._processInformation();
      gantt.render();
    }
    if (prevProps.listTitle !== this.props.listTitle) {
      //gantt.destructor();
      this._processInformation();
    }
  }

  private setZoom(value){
    switch (value){
      case 'Hours':
        gantt.config.scale_unit = 'day';
        gantt.config.date_scale = '%d %M';

        gantt.config.scale_height = 60;
        gantt.config.min_column_width = 30;
        gantt.config.subscales = [
          {unit:'hour', step:1, date:'%H'}
        ];
        break;
      case 'Days':
        gantt.config.min_column_width = 70;
        gantt.config.scale_unit = "week";
        gantt.config.date_scale = "#%W";
        gantt.config.subscales = [
          {unit: "day", step: 1, date: "%d %M"}
        ];
        gantt.config.scale_height = 60;
        break;
      case 'Months':
        gantt.config.min_column_width = 70;
        gantt.config.scale_unit = "month";
        gantt.config.date_scale = "%F";
        gantt.config.scale_height = 60;
        gantt.config.subscales = [
          {unit:"week", step:1, date:"#%W"}
        ];
        break;
      default:
        break;
    }
  }

  private initGanttEvents(): void {
    if(gantt.ganttEventsInitialized) {
      return;
    }
    gantt.ganttEventsInitialized = true;

    gantt.attachEvent('onAfterTaskAdd', (id, task) => {

      if ( task.users.length > 0 && task.sendmail ) {
        pnp.sp.web.ensureUser(task.users).then( res => {
          let users: number[] = [res.data.Id];
          const emailProps: EmailProperties = {
            To: [res.data.Email],
            Subject: "En ny oppgave er tilordnet deg",
            Body: `
              <div>
                Oppgaven <strong>${task.text}</strong> ble opprettet ${moment().format("DD/MM [kl] HH:mm:ss")}
              </div>
              <div>
                <p>
                  <strong>Oppgavetekst:</strong>
                </p>
                <p>
                  ${task.Body}
                </p>
              </div>
              <div>
                <p>
                  <strong>Start og slutt:</strong>
                </p>
                <p>
                  Start: ${moment(task.start_date).format("DD/MM/YYYY")}
                </p>
                <p>
                  Slutt: ${moment(task.end_date).format("DD/MM/YYYY")}
                </p>
              </div>
              <div>
                Se alle oppgaver her: <a href=${window.location.href}>Oppgaveliste</a>
              </div>
            `
          };
          const updatedTask: IGanttChartItemProp = {
            Title: task.text,
            Body: task.Body,
            StartDate: moment(task.start_date).format("YYYY-MM-DDTHH:MM:ssZ"),
            DueDate: moment(task.end_date).format("YYYY-MM-DDTHH:MM:ssZ"),
            PercentComplete: task.progress,
            AssignedToId: {
              results: users
            },
            PredecessorsId: {
              results: [task.parent] || []
            }
          };

          pnp.sp.web.lists.getByTitle(this.props.listTitle).items.add(updatedTask).then( res => {
            gantt.changeTaskId( task.id, res.data.ID );
            gantt.render();
            pnp.sp.utility.sendEmail(emailProps).then( mail => {
                //console.log("Email Sent!");
            });
          }).catch( e => {
            gantt.alert({
              title:"Error updating task",
              type:"alert-error",
              text: JSON.stringify(e)
            });
          });
        });
      } else {
        let users: number[] = [];

          const updatedTask: IGanttChartItemProp = {
            Title: task.text,
            Body: task.Body,
            StartDate: moment(task.start_date).format("YYYY-MM-DDTHH:MM:ssZ"),
            DueDate: moment(task.end_date).format("YYYY-MM-DDTHH:MM:ssZ"),
            PercentComplete: task.progress,
            AssignedToId: {
              results: users
            },
            PredecessorsId: {
              results: [task.parent] || []
            }
          };

          pnp.sp.web.lists.getByTitle(this.props.listTitle).items.add(updatedTask).then( res => {
            gantt.changeTaskId( task.id, res.data.ID );
            gantt.render();
          }).catch( e => {
            gantt.alert({
              title:"Error updating task",
              type:"alert-error",
              text: JSON.stringify(e)
            });
          });
      }
    });

    gantt.attachEvent('onAfterTaskUpdate', (id, task) => {
      // console.log("onAfterTaskUpdate");
      // console.log(id);
      // console.dir(task);

      //let users: number[] = [];

      if (task.users.length > 0 && task.sendmail) {
        pnp.sp.web.ensureUser(task.users).then( res => {
          let users: number[] = [res.data.Id];

          const emailProps: EmailProperties = {
            To: [res.data.Email],
            Subject: "En oppgave du er tilordnet har blitt oppdatert",
            Body: `
              <div>
                Oppgaven <strong>${task.text}</strong> har blitt endret ${moment().format("DD/MM [kl] HH:mm:ss")}
              </div>
              <div>
                <p>
                  <strong>Oppgavetekst:</strong>
                </p>
                <p>
                  ${task.Body}
                </p>
              </div>
              <div>
                <p>
                  <strong>Start og slutt:</strong>
                </p>
                <p>
                  Start: ${moment(task.start_date).format("DD/MM/YYYY")}
                </p>
                <p>
                  Slutt: ${moment(task.end_date).format("DD/MM/YYYY")}
                </p>
              </div>
              <div>
                Se alle oppgaver her: <a href=${window.location.href}>Oppgaveliste</a>
              </div>
            `
          };

          const updatedTask: IGanttChartItemProp = {
            Title: task.text,
            Body: task.Body,
            StartDate: moment(task.start_date).format("YYYY-MM-DDTHH:MM:ssZ"),
            DueDate: moment(task.end_date).format("YYYY-MM-DDTHH:MM:ssZ"),
            PercentComplete: task.progress,
            AssignedToId: {
              results: users
            }
          };

          pnp.sp.web.lists.getByTitle(this.props.listTitle).items.getById(id).update(updatedTask).then( res => {
            pnp.sp.utility.sendEmail(emailProps).then( mail => {
                //console.log("Email Sent!");
            });
          }).catch( e => {
            gantt.alert({
              title:"Error updating task",
              type:"alert-error",
              text: JSON.stringify(e)
            });
          });
        });
      } else {
        let users: number[] = [];

          const updatedTask: IGanttChartItemProp = {
            Title: task.text,
            Body: task.Body,
            StartDate: moment(task.start_date).format("YYYY-MM-DDTHH:MM:ssZ"),
            DueDate: moment(task.end_date).format("YYYY-MM-DDTHH:MM:ssZ"),
            PercentComplete: task.progress,
            AssignedToId: {
              results: users
            }
          };

          pnp.sp.web.lists.getByTitle(this.props.listTitle).items.getById(id).update(updatedTask).then( res => {
            // console.log("successfully updated item " + id);
            // console.dir(res);
          }).catch( e => {
            gantt.alert({
              title:"Error updating task",
              type:"alert-error",
              text: JSON.stringify(e)
            });
          });
      }
    });

    gantt.attachEvent('onAfterTaskDelete', (id) => {
      // console.log("onAfterTaskDelete");
      // console.log(id);
      pnp.sp.web.lists.getByTitle(this.props.listTitle).items.getById(id).delete().then( res => {
        //console.log("successfully deleted " + id);
      }).catch( e => {
        gantt.alert({
          title:"Error deleting task",
          type:"alert-error",
          text: JSON.stringify(e)
        });
      });
    });

    gantt.attachEvent('onAfterLinkAdd', (id, link) => {
      // console.log("onAfterLinkAdd");
      // console.log(id);
      // console.dir(link);
      pnp.sp.web.lists.getByTitle(this.props.listTitle).items.getById(link.target).select("PredecessorsId, ID").get().then( item => {
        //console.dir(item);
        let parents: number[] = item.PredecessorsId;
        if ( parents.indexOf(link.source) == -1 ) { //check if the link already exists
          parents.push(link.source);
          const updatedItem: IGanttChartItemProp = {
            PredecessorsId: {
              results: parents
            }
          };
          pnp.sp.web.lists.getByTitle(this.props.listTitle).items.getById(link.target).update(updatedItem).then( res => {
            console.log("successfully updated item " + link.target);
            //console.dir(res);
            //gantt.setParent(link.target, link.source);
            //gantt.render();
          }).catch( e => {
            gantt.alert({
              title:"Error updating link",
              type:"alert-error",
              text: JSON.stringify(e)
            });
            console.error(JSON.stringify(e));
          });
        } else {
          // console.log("link already exists");
        }
      });
    });

    gantt.attachEvent('onAfterLinkUpdate', (id, link) => {
      // console.log("onAfterLinkUpdate");
      // console.log(id);
      // console.dir(link);
    });

    gantt.attachEvent('onAfterLinkDelete', (id, link) => {
      // console.log("onAfterLinkDelete");
      // console.log(id);
      // console.dir(link);
      pnp.sp.web.lists.getByTitle(this.props.listTitle).items.getById(link.target).select("PredecessorsId, ID").get().then( item => {
        // console.dir(item);
        let parents: number[] = item.PredecessorsId;
        let indexOfLinkToRemove = parents.indexOf(link.source);
        // console.log("removing id " + link.source + " with index " + indexOfLinkToRemove);
        if ( indexOfLinkToRemove > -1 ) { //check if the link already exists
          parents.splice(indexOfLinkToRemove, 1);
          const updatedItem: IGanttChartItemProp = {
            PredecessorsId: {
              results: parents
            }
          };
          pnp.sp.web.lists.getByTitle(this.props.listTitle).items.getById(link.target).update(updatedItem).then( res => {
            // console.log("successfully updated item " + link.target);
            gantt.render();
          }).catch( e => {
            gantt.alert({
              title:"Error updating link",
              type:"alert-error",
              text: JSON.stringify(e)
            });
            console.error(JSON.stringify(e));
          });
        } else {
          // console.log("link already exists");
        }
      });
    });

    gantt.form_blocks["task_description"] = {
      render: function (sns) {
        // console.dir(sns);
        // return "<div class='gantt_cal_ltext' style='height:60px;'><br/>Assigned to&nbsp;<input type='text'></div>";
        return `<div class='gantt_cal_ltext' style='height:200px;'>
                  Title<br />
                  <input type="text" style="width:100%"><br />
                  Details<br />
                  <textarea></textarea>
                </div>`;
      },
      set_value: function (node, value, task) {
        // console.log("set_value");
        // console.dir(value);
        // console.dir(node);
        // console.dir(task);
        node.childNodes[3].value = value || "";
        node.childNodes[8].value = task.Body || "";
        let editor = wysiwyg(node.childNodes[8]);
        editor.onUpdate( () => {
          task.Body = editor.read();
        });
      },
      get_value: function (node, task) {
        // console.log("get_value");
        // console.dir(node);
        // let editor = wysiwyg(node.childNodes[8])
        // console.dir(task);
        // task.Body = node.childNodes[8].value;
        // console.dir(task);
        return node.childNodes[3].value;
      },
      focus: function (node) {
        var a = node.childNodes[3];
        a.select();
        a.focus();
      }
    };

    gantt.locale.labels.section_assigned = "Assigned to";
    let combo = null;

    gantt.form_blocks["assigned_to"] = {
      render: function (sns) {
        //return "<div class='gantt_cal_ltext' style='height:60px;'><br/>Assigned to&nbsp;<input type='text'></div>";
        return `<div class='gantt_cal_ltext' style='height:100px;'>
                  <div id='combo_zone' style='width:370px; height:30px;'></div>
                  Send email update?<input type="checkbox"></input>
                </div>`;
      },
      set_value: function (node, value, task) {
        // console.log("set_value");
        // console.dir(node);
        // console.dir(task);
        if(combo){
          combo.unload();
          combo = null;
        }

        if ( task.hasOwnProperty("sendmail") ) {
          node.childNodes[3].checked = task.sendmail;
        } else {
          node.childNodes[3].checked = true;
        }

        combo = new dhtmlXCombo("combo_zone","combo",350);

        combo.enableFilteringMode(true, "custom");

        combo.attachEvent("onDynXLS", function (text, ind) {
          // console.log("fired onDynXLS");
          // console.log(text);
          var text_length = text.replace(/^\s{1,}/, "").replace(/\s{1,}$/, "").length;
          if (text_length == 0) {
            combo.closeAll();
            combo.clearAll();
          } else if (text_length >= 3) {
            let query = SearchQueryBuilder.create().text("*" + text + "*").sourceId("B09A7990-05EA-4AF9-81EF-EDFAB16C4E31");

            pnp.sp.search(query).then( (results: any) => {
              const options: ganttOptions = {
                options: []
              };
              const people: peopleSearchResults[] = [];
              for (let p of results.PrimarySearchResults) {
                if( p.WorkEmail ) {
                  people.push({
                    "value": p.WorkEmail,
                    "text": p.WorkEmail
                  });
                }
              }

              options.options = people;

              if (people.length > 0) {
                combo.clearAll();
                combo.load(options);
                combo.openSelect();
              }
            }).catch( e => {
              console.error(JSON.stringify(e));
            });
          }
        });


        combo.DOMelem_input.value = task.users || "";
      },
      get_value: function (node, task) {
        // console.log("get_value");
        // console.dir(node);
        // console.dir(task);
        task.sendmail = node.childNodes[3].checked;
        task.users = combo.DOMelem_input.value || "";
        return task.users;
      },
      focus: function (node) {
      }
    };

    gantt.config.lightbox.sections = [
      { name:"description", height:300, map_to:"text", type:"task_description", focus:true},
      { name:"assigned", height:200, map_to:"users", type:"assigned_to", focus:false},
      { name:"time", height:72, type:"duration", map_to:"auto"}
    ];

    gantt.config.columns = [
      {name:"text", label:"Task name", tree:true, min_width:200, max_width: 270 },
      {name:"start_date", align:"center", width: 90},
      {name:"duration", align:"center", width: 46},
      {name:"add", width:40}
  ];

    //suppress enter key = save and close modal (to allow multi line edit for task details)
    gantt.keys.edit_save = -1;
}

  private _processInformation() {
    this.loadData( "*, AssignedTo/EMail", this.props.listTitle ).then( res => {
      this.setState({
        loading: false
      }, function() {
        gantt.init("gantt_here");
        this.initGanttEvents();
        gantt.config.xml_date="%d/%m/%Y %H:%i";
        gantt.parse(res);
      });
    }).catch( e=> {
      gantt.alert({
        title:"Error fetching items from list",
        type:"alert-error",
        text: JSON.stringify(e)
      });
    });
  }

  private handleZoomChange(zoom): void {
    this.setState({
      zoom: zoom
    });
  }

  private loadData(select: string, listname:string) {
    return new Promise<any>((resolve,reject) => {
    if(Environment.type === EnvironmentType.Local) {
      const tasks: IGanttDataObject = {
        data: [
          {
            id: 1, text: "Project #2", start_date: moment().format("DD/MM/YYYY"), duration: 4, order: 10,
            progress: 0.4, open: true
          },
          {
            id: 2, text: "Task #1", start_date: moment().add(1,"days").format("DD/MM/YYYY"), duration: 2, order: 10,
            progress: 0.6, parent: 1
          },
          {
            id: 3, text: "Task #2", start_date: moment().add(2,"days").format("DD/MM/YYYY"), duration: 2, order: 20,
            progress: 0.6, parent: 1
          }
        ],
        links: [
          {id: 1, source: 1, target: 2, type: "1"},
          {id: 2, source: 2, target: 3, type: "0"}
        ]
      };
      resolve(tasks);
    } else {
        pnp.sp.web.lists.getByTitle(listname).items
        .select(select)
        .expand("AssignedTo")
        .top(1000)
        .get().then( res => {
          if (res.length > 0) {
            this.setState({
              height: res.length * this.rowHeight + 70
            })
          }
          const data: IGanttData[] = res.map( obj => {
            let users = "";
            if ( obj.hasOwnProperty("AssignedTo") ) {
              users = obj.AssignedTo[0].EMail;
            }
            let rObj: IGanttData = {
              id: obj.ID,
              text: obj.Title,
              Body: obj.Body,
              start_date: moment(obj.StartDate).format("DD/MM/YYYY"),
              end_date: moment(obj.DueDate).format("DD/MM/YYYY"),
              progress: obj.PercentComplete,
              parent: obj.PredecessorsId[obj.PredecessorsId.length-1],
              users: users || "",
              open: true,
              sendmail: true
            };
            return rObj;
          });
          let linkId: number = 0;

          const links: IGanttLink[] = [];
          for ( let obj of res ) {
            for ( let i in obj.PredecessorsId ) {
              let rObj: IGanttLink = {
                id: linkId,
                source: obj.PredecessorsId[i],
                target: obj.ID,
                type: "0"
              };
              linkId++;
              links.push(rObj);
            }
          }
          const links1: IGanttLink[] = res.map( obj => {
            if( obj.PredecessorsId.length > 0 ) {
              for ( let i in obj.PredecessorsId ) {
                let rObj: IGanttLink = {
                  id: linkId,
                  source: obj.PredecessorsId[i],
                  target: obj.ID,
                  type: "0"
                };
                linkId++;
                return rObj;
              }
            }
          });
          const ganttData: IGanttDataObject = {
            data: data,
            links: links
          };
          resolve(ganttData);
        }).catch( err => {
          reject(JSON.stringify(err));
        });
      }
    });
  }

  public render(): React.ReactElement<IGanttChartProps> {
    const divStyle = {
      width: '100%',
      height: this.state.height + 'px',
    };
    let view = <Spinner size={SpinnerSize.large} label='Laster data' />;
    let zoomRadios = ['Hours', 'Days', 'Months'].map((value) => {
      let isActive = this.state.zoom === value;
      return (
        <label key={value} className={`${styles["radio-label"]} ${isActive ? styles["radio-label-active"]: ''}`}>
          <input type='radio'
             checked={isActive}
             //onChange={this.handleZoomChange}
             onClick={() => this.handleZoomChange(value)}
             value={value}/>
          {value}
        </label>
      );
    });
    if( !this.props.listTitle ) {
      view = <div>
        <b>Please select a task list from the web part property dialog</b>
      </div>
    }
    if ( !this.state.loading && this.props.listTitle ) {
      this.setZoom(this.state.zoom);
      view = <div className={ styles.ganttChart }>
      <div className={styles["zoom-bar"]}>
        <b>Zooming: </b>
          {zoomRadios}
      </div>
      <div
        id={"gantt_here"}
        style={divStyle}>
      </div>
    </div>;
    }
    return (
      <div className={ styles.ganttChart }>
        {view}
      </div>
    );
  }
}
