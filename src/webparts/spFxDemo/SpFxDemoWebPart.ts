import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpFxDemoWebPart.module.scss';
import * as strings from 'SpFxDemoWebPartStrings';

//[2].导入获取Url参数的模块
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';

//[3].导入EnvironmentType模块
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

//[4].导入JQuery模块
import * as $ from 'jquery';

//[6].如何导出Excel
import * as exceljs from 'exceljs';
import saveAs from 'file-saver';

export interface ISpFxDemoWebPartProps {
  description: string;
}


export default class SpFxDemoWebPart extends BaseClientSideWebPart<ISpFxDemoWebPartProps> {
  //[5].添加单击事件
  private _setButtonEventHandlers(): void {
    const webPart: SpFxDemoWebPart = this;
    this.domElement.querySelector('#btnRegister').addEventListener('click', () => {
      alert("单击事件");
    });
  }

  public render(): void {
    //[1].使用_spPageContextInfo
    var webUrl = this.context.pageContext.web.absoluteUrl;
    var currentUser = this.context.pageContext.user.loginName;

    //[2].获取Url参数
    var queryParms = new UrlQueryParameterCollection(window.location.href);
    var myParm = queryParms.getValue("myParm");
    if (myParm == null)
      myParm = "null";

    //[3].使用EnvironmentType模块
    var cE = Environment.type;
    var cET = EnvironmentType[cE];

    this.domElement.innerHTML = `
      <div class="${ styles.spFxDemo}">
    <div class="${ styles.container}">
      <div class="${ styles.row}">`
      //[1].使用_spPageContextInfo
      + `<div class="${styles.column}">[1].使用_spPageContextInfo<br/>WebUrl: ${webUrl}<br/>CurrentUser: ${currentUser}</div>`

      //[2].获取Url参数
      + `<div class="${styles.column}">[2].获取Url参数<br/>myParm: ${myParm}</div>`

      //[3].使用EnvironmentType模块
      + `<div class="${styles.column}">[3].使用EnvironmentType模块<br/>cE: ${cE}<br/>cET: ${cET}</div>`

      //[4].添加Jquery CDN引用HTML
      + `<div class="${styles.column}">[4].添加Jquery CDN引用HTML<br/><span id='myContent'></span></div>`

      //[5].添加单击事件
      + `<div class="${styles.column}">[5].添加单击事件<br/><button type="button" id="btnRegister">Click Me!</button></div>`

      //[6].如何导出Excel
      + `<div class="${styles.column}">[6].如何导出Excel<br/><button type="button" id="btnExport">Export Excel!</button></div>`

      //     + `<div class="${styles.column}">br/
      //         <span class="${ styles.title}">Welcome to SharePoint!</span>
      // <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
      //   <p class="${ styles.description}">${escape(this.properties.description)}</p>
      //     <a href="https://aka.ms/spfx" class="${ styles.button}">
      //       <span class="${ styles.label}">Learn more</span>
      //         </a>
      //         </div>`
      + `
          </div>
          </div>
          </div>`;

    //[4].添加Jquery CDN引用JavaScript脚本
    $("#myContent").html("test for jquery");
    //[6].如何导出Excel
    $("#btnExport").on("click",downloadExcel);
    function downloadExcel() {
      var workbook = new exceljs.Workbook();
      var sheet = workbook.addWorksheet('My Sheet');
      sheet.columns = [
        { header: 'Name', key: 'name', width: 30 },
        { header: 'Title', key: 'title', width: 50 }
      ];
      var data = [{
        name: "Tom",
        title: "Project manager"
      }, {
        name: "John",
        title: "Developer"
      }];
      data.forEach(function (currentValue, index, arr) {
        sheet.addRow({
          "name": currentValue.name,
          "title": currentValue.title
        });
      });
      var fileName = "data";
      workbook.xlsx.writeBuffer().then(function (buffer) {
        saveAs(new Blob([buffer], {
          type: 'application/octet-stream'
        }), fileName + '.' + 'xlsx');
      });
    }

    //[5].添加单击事件
    this._setButtonEventHandlers();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
