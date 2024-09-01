import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnPcrudWebPart.module.scss';
import * as strings from 'PnPcrudWebPartStrings';
import { sp } from "@pnp/sp/presets/all";


export interface IPnPcrudWebPartProps {
  description: string;
}

export default class PnPcrudWebPart extends BaseClientSideWebPart<IPnPcrudWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.pnPcrud }">
<div>
  <table border="5" bgcolor="aqua">
    <tr>
      <td>Please Enter Software ID </td>
      <td><input type="text" id="txtID"/></td>
        <td><input type="submit" id="btnRead" value="Read Details"/></td>
    </tr>

    <tr>
      <td>Software Title</td>
      <td><input type="text" id="txtSoftwareTitle"></td>
    </tr>

    <tr>
      <td>Software Name</td>
     <td><input type="text" id="txtSoftwareName"></td>
    </tr>

    <tr>
      <td>Software Vendor</td>
      <td>
        <select name="" id="ddlSoftwareVendor">
          <option value="Sun">sun</option>
          <option value="Oracle">Oracle</option>
          <option value="Microsoft">Microsoft</option>
        </select>
      </td>
    </tr>

    <tr>
      <td>Software Version</td>
      <td><input type="text" id="txtSoftwareVersion"></td>
    </tr>

    <tr>
      <td>Software Description</td>
      <td><textarea name="" id="txtSoftwareDescription" cols="40" rows="5"></textarea></td>
    </tr>

    <tr>
      <td colspan="2" align="center">
      <input type="submit" id="btnSubmit" value="Insert Item"/>
      <input type="submit" id="btnUpdate" value="Update"/>
      <input type="submit" id="btnDelete" value="Delete"/>
      </td>
    </tr>

  </table>
  </div>
  <div id="divStatus"></div>

<h2>Get All List Items</h2>
<hr/>

<div id="spListData"></div>


      </div>`;

    this._bindEvents();
    this.readAllItems();
  }
  readAllItems() {
let html: string = "<table border='1' width='100%'style='bordercollapse: collapse;'>"
html += `<th>ID</th><th>Title</th><th>Software Vendor</th><th>Software Version</th><th>Software Name</th><th>Software Description</th>`

sp.web.lists.getByTitle('SoftwareCatalog').items.get().then((items: any[])=>{
  items.forEach(function (item) {
    html += `<tr>

    <td>${item["ID"]}</td>
    <td>${item["Title"]}</td>
    <td>${item["SoftwareVendor"]}</td>
    <td>${item["SoftwareVersion"]}</td>
    <td>${item["SoftwareName"]}</td>
    <td>${item["SoftwareDescription"]}</td>
    </tr>`;

  });

html += `</table>`;
const allitems: Element = this.domElement.querySelector('#spListData');
allitems.innerHTML = html;
});
  }



  private _bindEvents(): void {
    this.domElement.querySelector('#btnSubmit').addEventListener('click', ()=> {this.addListItem();});
    this.domElement.querySelector('#btnRead').addEventListener('click', ()=>{this.readListItem();});
    this.domElement.querySelector('#btnUpdate').addEventListener('click', ()=>{this.updateListItem();});
    this.domElement.querySelector('#btnDelete').addEventListener('click', ()=>{this.deleteListItem();});
  }
 private deleteListItem():void {
const id = document.getElementById('txtID')['value'];
sp.web.lists.getByTitle('SoftwareCatalog').items.getById(id).delete().then(r => {
  alert('Item Deleted');
}).catch(e => {
  alert('Error: ' + e.message);
});
  }
  private updateListItem():void {
var title = document.getElementById('txtSoftwareTitle')['value'];
var softwareVender = document.getElementById('ddlSoftwareVendor')['value'];
var softwareVersion = document.getElementById('txtSoftwareVersion')['value'];
var softwareName = document.getElementById('txtSoftwareName')['value'];
var softwareDescription = document.getElementById('txtSoftwareDescription')['value'];

var id = document.getElementById('txtID')['value'];

sp.web.lists.getByTitle('SoftwareCatalog').items.getById(id).update({
  Title: title,
  SoftwareVendor: softwareVender,
  SoftwareVersion: softwareVersion,
  SoftwareName: softwareName,
  SoftwareDescription: softwareDescription
}).then(r => {
  alert('Item Updated');
}).catch(e => {
  alert('Error: ' + e.message);
});
  }
  private readListItem() :void {

    const id=document.getElementById('txtID')['value'];
    // const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('SoftwareCatalog')/items("+id+")";

    sp.web.lists.getByTitle('SoftwareCatalog').items.getById(id).get().then(r => {
      document.getElementById('txtSoftwareTitle')['value']=r.Title;
      document.getElementById('txtSoftwareName')['value']=r.SoftwareName;
      document.getElementById('txtSoftwareVersion')['value']=r.SoftwareVersion;
      document.getElementById('ddlSoftwareVendor')['value']=r.SoftwareVendor;
      document.getElementById('txtSoftwareDescription')['value']=r.SoftwareDescription;
    }).catch(e => {
      alert('Error: ' + e.message);
    });


  }
  private addListItem(): void {
    var softwaretitle = document.getElementById('txtSoftwareTitle')['value']
    var softwarename = document.getElementById('txtSoftwareName')['value']
    var softwareversion = document.getElementById('txtSoftwareVersion')['value']
    var softwarevendor = document.getElementById('ddlSoftwareVendor')['value']
    var softwareDescription = document.getElementById('txtSoftwareDescription')['value']

    // const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('SoftwareCatalog')/items";

    sp.web.lists.getByTitle('SoftwareCatalog').items.add({
      Title: softwaretitle,
      SoftwareName: softwarename,
      SoftwareVersion: softwareversion,
      SoftwareVendor: softwarevendor,
      SoftwareDescription: softwareDescription
    }).then(r => {
      alert('Item Added');
    }).catch(e => {
      alert('Error: ' + e.message);
    });
  }





  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

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
