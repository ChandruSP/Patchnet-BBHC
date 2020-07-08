import * as React from 'react';

import styles from "./Providers.module.scss";
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';

import { Announced } from 'office-ui-fabric-react/lib/Announced';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { Link } from 'office-ui-fabric-react/lib/Link';

import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';

import { Label } from "office-ui-fabric-react/lib/Label";
import { Image, IImageProps } from "office-ui-fabric-react/lib/Image";
import { CommandBar, ICommandBarStyles } from 'office-ui-fabric-react/lib/CommandBar';

import { ExcelRenderer } from "react-excel-renderer";
import { useId, useBoolean } from '@uifabric/react-hooks';

import { getId } from "office-ui-fabric-react/lib/Utilities";
import {
  IStackTokens,
  Stack,
  IStackProps,
  IStackStyles,
} from "office-ui-fabric-react/lib/Stack";
import * as ReactIcons from "@fluentui/react-icons";
import { mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
import { IconButton } from "@fluentui/react/lib/Button";


import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';

import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";
import { PermissionKind } from "@pnp/sp/security";

import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
} from "office-ui-fabric-react/lib/Dropdown";

import "alertifyjs";

import "../../../ExternalRef/CSS/style.css";
import "../../../ExternalRef/CSS/alertify.min.css";
var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");

import { IProvidersProp } from './IProvidersProps';

const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
});

const textFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };


const currentYear = new Date().getFullYear();
const fileId = getId("anInput");

export interface ILocalProvidersProp {
  Id: number;
  Title: string;
  LegalName: string;
  ProviderID: string;
  TemplateType: string;
  ContractId: string;
  Users: string;
  IsDeleted: boolean;
  Logs: string;
}

export interface IDetailsListBasicExampleState {
  items: ILocalProvidersProp[];
  allItems: ILocalProvidersProp[];
  selectionDetails: string;
  hideDialog: boolean;
  hideDeleteDialog: boolean;

  editUsers: string;
  providerName: string;
  folders: IDropdownOption[];
  subFolders: IDropdownOption[];
  cols: [];
  rows: [];
  AllUsers: any[];
  formData: {
    ProviderID: string;
    Title: string;
    LegalName: string;
    ContractId: string;
    TemplateType: string;
    Users: string;
    Id: number;
    IsDeleted: boolean;
    Logs: string;
  };
  fileName: "";
}



export default class Providers extends React.Component<IProvidersProp, IDetailsListBasicExampleState> {

  private _selection: Selection;
  private _columns: IColumn[];

  selUsers = [];
  allUsers = [];
  fileObj = null;
  rootFolder = "Providers Library";
  templateTypes = [{
    key: "Contract Providers",
    text: "Contract Provider",
  }, {
    key: "Agreement Providers",
    text: "Agreement Provider",
  }];

  contributePermission = null;
  readPermission = null;
  currentUser = null;
  userDetails = [];

  constructor(props) {
    super(props);

    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      },
    });

    alertify.set("notifier", "position", "top-right");

    var that = this;
    sp.web.roleDefinitions.getByName('Read').get().then(function (res) {
      that.readPermission = res.Id;
    });

    sp.web.roleDefinitions.getByName('Contribute').get().then(function (res) {
      that.contributePermission = res.Id;
    });

    this.currentUser = sp.web.currentUser();

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    this._columns = [
      // { key: 'column1', name: 'Id', fieldName: 'Id', minWidth: 100, maxWidth: 200, isResizable: true },
      {
        key: 'column1', name: 'Provider Name', fieldName: 'Title', minWidth: 100, maxWidth: 200, isResizable: true, isSorted: true,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick
      },
      {
        key: 'column2', name: 'Legal Name', fieldName: 'LegalName', minWidth: 100, maxWidth: 200, isResizable: true, isSorted: true,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick
      },
      { key: 'column3', name: 'Provider ID', fieldName: 'ProviderID', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column4', name: 'Template Type', fieldName: 'TemplateType', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column5', name: 'Contract Id', fieldName: 'ContractId', minWidth: 100, maxWidth: 200, isResizable: true },
    ];

    that.state = {
      items: [],
      allItems: [],
      selectionDetails: '',
      hideDialog: true,
      hideDeleteDialog: true,
      editUsers: "",

      providerName: "",
      folders: [],
      subFolders: [],
      cols: [],
      rows: [],
      AllUsers: [""],
      formData: {
        ProviderID: "",
        Title: "",
        LegalName: "",
        ContractId: "",
        TemplateType: "Contract Providers",
        Users: "",
        Id: 0,
        IsDeleted: false,
        Logs: ''
      },
      fileName: "",
    };
    this.loadTableData();
  }

  loadTableData() {
    var that = this;
    sp.web.lists
      .getByTitle("ProviderDetails")
      .items.select("Title", "LegalName", "ProviderID", "TemplateType", "ContractId", "Id", "Users", "IsDeleted", "Logs").get().then(function (data) {
        var allItems = that.state.items;
        allItems = [];
        for (let index = 0; index < data.length; index++) {
          const element = data[index];
          if (!element.IsDeleted) {
            allItems.push({
              Id: element.Id,
              Title: element.Title,
              LegalName: element.LegalName,
              ProviderID: element.ProviderID,
              TemplateType: element.TemplateType,
              ContractId: element.ContractId,
              Users: element.Users,
              Logs: element.Logs,
              IsDeleted: element.IsDeleted
            });
          }
        }
        that.setState({ items: allItems, allItems: allItems, selectionDetails: that._getSelectionDetails() });
      });
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected';
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
      items: text ? this.state.allItems.filter(i => i.Title.toLowerCase().indexOf(text) > -1) : this.state.allItems,
    });
  };

  private _onItemInvoked = (item: ILocalProvidersProp): void => {
    alert(`Item invoked: ${item.Title}`);
  };

  private editItem(element) {
    var id = element.target.id;
    var data = this.state.allItems.filter(c => c.Id == id);
    var formData = this.state.formData;
    var allUsers = this.state.AllUsers;
    formData = {
      Id: parseInt(id),
      Title: data[0].Title,
      ProviderID: data[0].ProviderID,
      LegalName: data[0].LegalName,
      ContractId: data[0].ContractId,
      TemplateType: data[0].TemplateType,
      Users: "",
      IsDeleted: false,
      Logs: data[0].Logs
    };
    allUsers = [];
    var susers = data[0].Users.split(';');
    for (let index = 0; index < susers.length; index++) {
      if (susers[index]) {
        allUsers.push(susers[index]);
      }
    }
    this.setState({ editUsers: data[0].Users, AllUsers: allUsers, formData: formData, hideDialog: false });
  }

  hideDialog() {
    this.setState({ hideDialog: true });
  }

  providerNameChange = (event) => {
    this.setState({ providerName: event.target.value });
    this.inputChangeHandler(event);
  };

  inputChangeHandler(e) {
    let formData = this.state.formData;
    formData[e.target.name] = e.target.value;
    this.setState({
      formData,
    });
  }

  processInputProvider = () => {
    var formData = this.state.formData;
    if (!formData.ProviderID) {
      alertify.error("Provider ID is required");
      return;
    }
    if (!formData.Title) {
      alertify.error("Provider name is required");
      return;
    }
    if (!formData.ContractId) {
      alertify.error("Contract Id is required");
      return;
    }
    if (!formData.LegalName) {
      alertify.error("Legal name is required");
      return;
    }
    var that = this;
    for (let index = 0; index < this.state.AllUsers.length; index++) {
      const user = this.state.AllUsers[index];
      if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(user)) {
        formData.Users = formData.Users + user + ";";
        sp.web.siteUsers.getByEmail(user).get().then(function (data) {
          that.userDetails.push({
            Id: data.Id,
            Email: user
          });
        });
      } else {
        alertify.error("User " + (index + 1) + " not valid");
        return;
      }
    }

    if (this.state.editUsers) {
      var existingUsers = this.state.editUsers.split(';');
      var newUsers = this.state.formData.Users.split(';');

      for (let index = 0; index < newUsers.length; index++) {
        if (newUsers[index]) {
          var exist = existingUsers.filter(c => c == newUsers[index]);
          if (exist.length == 0) {
            that.setpermissionfornewuser("TemplateLibrary/" + that.state.formData.TemplateType, newUsers[index], true);
            that.setpermissionformaintemplate("TemplateLibrary/" + this.state.formData.TemplateType, newUsers[index]);
          }
        }
      }

      for (let j = 0; j < existingUsers.length; j++) {
        if (existingUsers[j]) {
          var removeuser = newUsers.filter(c => c == existingUsers[j]);
          if (removeuser.length == 0) {
            that.setpermissionfornewuser("TemplateLibrary/" + that.state.formData.TemplateType, existingUsers[j], false);
          }
        }
      }
    } else {
      for (let index = 0; index < this.state.AllUsers.length; index++) {
        const user = this.state.AllUsers[index];
        if (user) {
          that.setpermissionformaintemplate("TemplateLibrary/" + this.state.formData.TemplateType, user);
        }
      }
    }

    this.setState({ formData: formData });
    this.addToList(currentYear, this.state.formData);
  };


  setpermissionfornewuser(folderPath, user, addpermission) {
    var reacthandler = this;
    sp.web
      .getFolderByServerRelativePath(folderPath)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          reacthandler.setpermission(0, data, user, addpermission);
        }
      });
  }


  setpermission(index, data, user, addpermission) {
    var reacthandler = this;
    var clonedUrl = data[index].ServerRelativeUrl;
    var url = clonedUrl.replace(this.props.currentContext.pageContext.web.serverRelativeUrl + '/', '');
    const spHttpClient: SPHttpClient = this.props.currentContext.spHttpClient;

    var contract = this.state.formData.ContractId.substr(this.state.formData.ContractId.length - 2, 2);
    var nextyear = parseInt(contract) + 1;
    var currentyearprefix = currentYear.toString().substr(0, 2);
    var yearfolder = "FY " + (currentyearprefix + contract) + "-" + (currentyearprefix + nextyear) + '/' + this.state.formData.Title;

    var providerFolder = reacthandler.rootFolder + "/" + yearfolder;
    var mainTemplateFolder = "TemplateLibrary/" + reacthandler.state.formData.TemplateType
    url = url.replace(mainTemplateFolder, providerFolder);

    url = url.replace(' - Upload', '');

    var queryUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativeUrl(" + "'" + url + "'" + ")/ListItemAllFields/breakroleinheritance(false)";
    const spOpts: ISPHttpClientOptions = {};

    sp.web.siteUsers.getByEmail(user).get().then(function (userdata) {
      spHttpClient.post(queryUrl, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
        if (response.ok) {
          var permission = reacthandler.readPermission;
          var sdata = clonedUrl.split('/');
          if (sdata[sdata.length - 1].toLocaleLowerCase().indexOf('upload') > 0) {
            permission = reacthandler.contributePermission;
          }

          var postUrl = reacthandler.props.currentContext.pageContext.web.absoluteUrl + '/_api/web/GetFolderByServerRelativeUrl(' + "'" + url + "'" + ')/ListItemAllFields/roleassignments/removeroleassignment(principalid=' + userdata.Id + ',roledefid=' + permission + ')';
          if (addpermission) {
            postUrl = reacthandler.props.currentContext.pageContext.web.absoluteUrl + '/_api/web/GetFolderByServerRelativeUrl(' + "'" + url + "'" + ')/ListItemAllFields/roleassignments/addroleassignment(principalid=' + userdata.Id + ',roledefid=' + permission + ')';
          }
          spHttpClient.post(postUrl, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
            if (response.ok) {
            }
          });


        }
      });
    });
    reacthandler.setpermissionfornewuser(data[index].ServerRelativeUrl, user, addpermission);
    index = index + 1;
    if (index < data.length) {
      reacthandler.setpermission(index, data, user, addpermission);
    }
  }

  setpermissionformaintemplate(url, useremail) {
    var that = this;
    sp.web
      .getFolderByServerRelativePath(url)
      .folders.get()
      .then(function (data) {
        for (let index = 0; index < data.length; index++) {
          that.setmaintemplatepermission(data[index].ServerRelativeUrl, useremail);
          that.setpermissionformaintemplate(data[index].ServerRelativeUrl, useremail);
        }
      });
  }

  setmaintemplatepermission(mainurl, usermail) {
    var reacthandler = this;
    var url = mainurl.replace(this.props.currentContext.pageContext.web.serverRelativeUrl + '/', '');
    const spHttpClient: SPHttpClient = this.props.currentContext.spHttpClient;
    var queryUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativeUrl(" + "'" + url + "'" + ")/ListItemAllFields/breakroleinheritance(false)";
    const spOpts: ISPHttpClientOptions = {};
    spHttpClient.post(queryUrl, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        var permission = reacthandler.readPermission;
        sp.web.siteUsers.getByEmail(usermail).get().then(function (userdata) {
          var postUrl = reacthandler.props.currentContext.pageContext.web.absoluteUrl + '/_api/web/GetFolderByServerRelativeUrl(' + "'" + url + "'" + ')/ListItemAllFields/roleassignments/addroleassignment(principalid=' + userdata.Id + ',roledefid=' + permission + ')';
          spHttpClient.post(postUrl, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
            if (response.ok) {
            }
          });
        });
      }
    })
  }

  addToList(year, formData) {
    var that = this;
    if (formData.Id > 0) {
      formData.Logs = formData.Logs + "\n\nUpdated on : " + new Date() + "\nUpdated by : " + this.props.currentContext.pageContext.user.displayName;
      sp.web.lists
        .getByTitle("ProviderDetails")
        .items.getById(formData.Id).update(formData)
        .then((res) => {
          alertify.success("Provider updated");
          that.loadTableData();
          that.setState({ hideDialog: true });
        });
    } else {
      var currentMonth = new Date().getMonth() + 1;
      var stryear = currentYear.toString().substr(2, 2);
      if (currentMonth < 7) {
        stryear = (currentYear - 1).toString().substr(2, 2);
      }
      formData.ContractId = formData.ContractId + '-' + stryear;
      formData.Logs = "Added on : " + new Date() + "\nAdded by : " + this.props.currentContext.pageContext.user.displayName;
      sp.web.lists
        .getByTitle("ProviderDetails")
        .items.add(formData)
        .then((res) => {
          that.createProvider(formData.Title, year, formData);
          that.loadTableData();
        });
    }
  }


  createProvider = (providerName, year, formData) => {
    var reacthandler = this;
    var currentMonth = new Date().getMonth() + 1;
    var stryear = year + "-" + (year + 1);
    if (currentMonth < 7) {
      stryear = (year - 1) + "-" + year;
    }

    var folderName = reacthandler.rootFolder + "/" + "FY " + stryear;

    sp.web.folders.add(folderName + "/" + providerName).then(function (data) {
      reacthandler.getFolder("TemplateLibrary/" + formData.TemplateType, providerName, year, formData);
    });
    alertify.success("Provider is created");
    reacthandler.setState({ hideDialog: true });
  };

  getFolder = (folderPath, providerName, year, formData) => {
    var reacthandler = this;
    sp.web
      .getFolderByServerRelativePath(folderPath)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          reacthandler.processFolder(0, data, providerName, year, formData);
        }
      });
  };

  async processFolder(index, data, providerName, year, formData) {
    var reacthandler = this;
    var currentMonth = new Date().getMonth() + 1;
    var stryear = year + "-" + (year + 1);
    if (currentMonth < 7) {
      stryear = (year - 1) + "-" + year;
    }
    var folderName = reacthandler.rootFolder + "/" + "FY " + stryear;
    var clonedUrl = data[index].ServerRelativeUrl.replace(
      "TemplateLibrary/" + formData.TemplateType,
      folderName + "/" + providerName
    );

    clonedUrl = clonedUrl.replace(' - Upload', '');

    // reacthandler.createFolder(clonedUrl);
    sp.web.folders.add(clonedUrl).then((res) => {
      var url = clonedUrl.replace(this.props.currentContext.pageContext.web.serverRelativeUrl + '/', '');
      const spHttpClient: SPHttpClient = this.props.currentContext.spHttpClient;
      var queryUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativeUrl(" + "'" + url + "'" + ")/ListItemAllFields/breakroleinheritance(false)";
      const spOpts: ISPHttpClientOptions = {};
      spHttpClient.post(queryUrl, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
        if (response.ok) {
          var permission = reacthandler.readPermission;
          var sdata = data[index].ServerRelativeUrl.split('/');
          if (sdata[sdata.length - 1].toLocaleLowerCase().indexOf('upload') > 0) {
            permission = reacthandler.contributePermission;
          }
          for (let index = 0; index < reacthandler.userDetails.length; index++) {
            const userData = reacthandler.userDetails[index];

            // var getPermissionUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativeUrl(" + "'" + data[index].ServerRelativeUrl + "'" + ")/ListItemAllFields/roleassignments?$expand=Member,RoleDefinitionBindings";


            // const requestHeaders: Headers = new Headers();
            // requestHeaders.append('Accept', 'application/json');
            // requestHeaders.append('Content-type', 'application/json');

            // const opt: ISPHttpClientOptions = { headers: requestHeaders };

            // spHttpClient.get(getPermissionUrl, SPHttpClient.configurations.v1, opt).then(function (permissions) {
            //   debugger;
            // });

            sp.web.getFolderByServerRelativeUrl(data[index].ServerRelativeUrl).expand("ListItemAllFields/RoleAssignments/Member", "ListItemAllFields/RoleAssignments/RoleDefinitionBindings", "ListItemAllFields/RoleAssignments/Member/Users").get().then((resdata) => {
              var roleAssignments = resdata["ListItemAllFields"].RoleAssignments;
              for (let i = 0; i < roleAssignments.length; i++) {
                const role = roleAssignments[i];
                for (let j = 0; j < role.RoleDefinitionBindings.length; j++) {
                  const definition = role.RoleDefinitionBindings[j];
                  var bbhcpostUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/_api/web/GetFolderByServerRelativeUrl(' + "'" + url + "'" + ')/ListItemAllFields/roleassignments/addroleassignment(principalid=' + role.Member.Id + ',roledefid=' + definition.Id + ')';
                  spHttpClient.post(bbhcpostUrl, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
                    if (response.ok) {
                    }
                  });
                }
              }
            });

            var postUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/_api/web/GetFolderByServerRelativeUrl(' + "'" + url + "'" + ')/ListItemAllFields/roleassignments/addroleassignment(principalid=' + userData.Id + ',roledefid=' + permission + ')';

            spHttpClient.post(postUrl, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
              if (response.ok) {
              }
            });

          }

        }
      });

      reacthandler.getFolder(data[index].ServerRelativeUrl, providerName, year, formData);
      index = index + 1;
      if (index < data.length) {
        reacthandler.processFolder(index, data, providerName, year, formData);
      }
    });
  }

  createFolder = async (folderPath) => {
    await sp.web.folders.add(folderPath);
  };

  uploadFile = async (event) => {
    var reacthandler = this;
    if (event.target.files && event.target.files.length > 0) {
      reacthandler.fileObj = event.target.files[0];
      reacthandler.setState({ fileName: event.target.files[0].name });
    } else {
      reacthandler.fileObj = null;
    }
  };


  userchange(event) {
    var allusers = this.state.AllUsers;
    allusers[parseInt(event.target.id)] = event.target.value;
    this.setState({ AllUsers: allusers });
  }

  removeuser(index) {
    var allusers = this.state.AllUsers;
    allusers.splice(index, 1);
    this.setState({ AllUsers: allusers });
  }

  newuser() {
    var allusers = this.state.AllUsers;
    allusers.push("");
    this.setState({ AllUsers: allusers });
  }

  templateChange(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void {
    var formData = this.state.formData;
    formData.TemplateType = option.key;
    this.setState({ formData: formData });
  }

  _onAddRow() {
    var formData = this.state.formData;
    var allUsers = this.state.AllUsers;
    formData = {
      Id: 0,
      Title: "",
      ProviderID: "",
      LegalName: "",
      ContractId: "",
      TemplateType: "Contract Providers",
      Users: "",
      IsDeleted: false,
      Logs: ""
    };
    allUsers = [""];
    this.setState({ editUsers: "", AllUsers: allUsers, formData: formData, hideDialog: false });
  }

  _onDeleteRow() {
    this.setState({ hideDeleteDialog: false })
  }

  hideDelete() {
    this.setState({ hideDeleteDialog: true })
  }

  deleteItems() {
    var selItems = this._selection.getSelection();
    if (selItems.length > 0) {
      this.updateDeleteTag(0, selItems);
    }
  }

  updateDeleteTag(index, items) {
    var that = this;
    var formData = items[index];
    formData.IsDeleted = true;
    formData.Logs = formData.Logs + "\n\nDeleted on : " + new Date() + "\nDeleted by : " + that.props.currentContext.pageContext.user.displayName;
    sp.web.lists
      .getByTitle("ProviderDetails")
      .items.getById(formData.Id).update(formData)
      .then((res) => {

        var users = formData.Users.split(';');
        for (let j = 0; j < users.length; j++) {
          const user = users[j];
          if (user) {
            that.setpermissionfornewuser("TemplateLibrary/" + that.state.formData.TemplateType, user, false);
          }
        }

        index = index + 1;
        if (index < items.length) {
          that.updateDeleteTag(index, items);
        } else {
          alertify.success("Provider deleted successfully");
          that.loadTableData();
          that.setState({ hideDeleteDialog: true });
        }
      });
  }



  _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const newColumns: IColumn[] = this._columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;

      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    var items = this.state.items;
    const newItems = this._copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      items: newItems,
    });
  }


  _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }

  public render(): React.ReactElement<IProvidersProp> {

    const _renderItemColumn = (item, index: number, column: IColumn) => {
      const fieldContent = item[column.fieldName] as string;
      switch (column.fieldName) {

        case 'Title':
          return (
            <Link id={item["Id"] + ''} onClick={this.editItem.bind(this)}>{fieldContent}</Link>
          );

        default:
          return <span>{fieldContent}</span>;
      }
    }


    const modelProps = {
      isBlocking: true,
      topOffsetFixed: true,
    };


    const stackTokens: IStackTokens = {
      childrenGap: 4,
    };
    const stackStyles: Partial<IStackStyles> = {
      root: {
        width: 600,
      },
    };

    const columnstyle: Partial<IStackProps> = {
      tokens: {
        childrenGap: 5,
      },
      styles: {
        root: {
          width: 300,
          paddingTop: 10,
        },
      },
    };

    const iconcolumnstyle: Partial<IStackProps> = {
      tokens: {
        childrenGap: 5,
      },
      styles: {
        root: {
          width: 30,
          paddingTop: 28,
        },
      },
    };


    const btnstackTokens: IStackTokens = {
      childrenGap: 4,
    };
    const btnstackStyles: Partial<IStackStyles> = {
      root: {
        width: 600,
      },
    };

    const btncolumnstyle: Partial<IStackProps> = {
      tokens: {
        childrenGap: 4,
      },
      styles: {
        root: {
          width: 100,
          paddingTop: 10,
        },
      },
    };

    const classes = mergeStyleSets({
      cell: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        margin: "80px",
        float: "left",
        height: "50px",
        width: "50px",
      },
      icon: {
        fontSize: "50px",
      },
      code: {
        background: "#f2f2f2",
        borderRadius: "4px",
        padding: "4px",
      },
      navigationText: {
        width: 100,
        margin: "0 5px",
      },
    });


    const commandBarStyles: Partial<ICommandBarStyles> = { root: { marginBottom: '40px' } };

    // const labelId: string = useId('dialogLabel');
    // const subTextId: string = useId('subTextLabel');
    // const dialogStyles = { main: { maxWidth: 450 } };
    // const [isDraggable, { toggle: toggleIsDraggable }] = useBoolean(false);

    // const modalProps = React.useMemo(
    //   () => ({
    //     titleAriaId: labelId,
    //     subtitleAriaId: subTextId,
    //     isBlocking: false,
    //     styles: dialogStyles,
    //     dragOptions: undefined,
    //   }),
    //   [isDraggable],
    // );

    const dialogContentProps = {
      type: DialogType.normal,
      title: 'Delete',
      closeButtonAriaLabel: 'Close',
      subText: 'Do you want to delete the selected Provider(s)?',
    };

    return (
      <div>

        <CommandBar
          styles={commandBarStyles}
          items={[
            {
              key: 'addRow',
              text: 'Add',
              iconProps: { iconName: 'Add' },
              onClick: this._onAddRow.bind(this),
            },
            {
              key: 'deleteRow',
              text: 'Delete row',
              iconProps: { iconName: 'Delete' },
              onClick: this._onDeleteRow.bind(this),
            }]}
        />

        <Fabric>
          <div className={exampleChildClass}>{this.state.selectionDetails}</div>
          <Announced message={this.state.selectionDetails} />
          <TextField
            className={exampleChildClass}
            label="Filter by provider name:"
            onChange={this._onFilter.bind(this)}
            styles={textFieldStyles}
          />
          <Announced message={`Number of items after filter applied: ${this.state.items.length}.`} />
          <MarqueeSelection selection={this._selection}>
            <DetailsList
              items={this.state.items}
              columns={this._columns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selection={this._selection}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="Row checkbox"
              onItemInvoked={this._onItemInvoked}
              onRenderItemColumn={_renderItemColumn}
            />
          </MarqueeSelection>
        </Fabric>


        <Dialog hidden={this.state.hideDialog} modalProps={modelProps} minWidth="400px">

          <Stack {...columnstyle}>

            {
              this.state.formData.Id == 0 ? <div><ChoiceGroup defaultSelectedKey={this.state.formData.TemplateType} options={this.templateTypes} onChange={this.templateChange.bind(this)} label="Provider Type" />

                <TextField
                  label="Provider ID"
                  onChange={(e) => this.inputChangeHandler.call(this, e)}
                  width="100px"
                  name="ProviderID"
                  value={this.state.formData.ProviderID}
                  required
                ></TextField>

                <TextField
                  label="Provider Name"
                  onChange={this.providerNameChange}
                  width="100px"
                  name="Title"
                  value={this.state.formData.Title}
                  required
                ></TextField>

                <TextField
                  label="Contract ID"
                  width="200px"
                  onChange={(e) => this.inputChangeHandler.call(this, e)}
                  value={this.state.formData.ContractId}
                  name="ContractId"
                  required
                ></TextField>

                <TextField
                  label="Legal Name"
                  width="200px"
                  onChange={(e) => this.inputChangeHandler.call(this, e)}
                  value={this.state.formData.LegalName}
                  name="LegalName"
                  required
                ></TextField>
              </div>
                : ''
            }

            {this.state.AllUsers.map((user, index) => {
              if (this.state.AllUsers.length == 1) {
                return (
                  <div>
                    <Stack
                      horizontal
                      tokens={stackTokens}
                      styles={stackStyles}
                    >
                      <Stack {...columnstyle}>
                        <TextField
                          label="User"
                          width="200px"
                          id={index + ""}
                          onChange={(e) => this.userchange.call(this, e)}
                          value={user}
                          name="userName"
                          required
                        ></TextField>
                      </Stack>

                      <Stack {...iconcolumnstyle}>
                        <IconButton
                          iconProps={{ iconName: "Add" }}
                          title="Add User"
                          ariaLabel="Add"
                          onClick={this.newuser.bind(this)}
                        />
                      </Stack>
                    </Stack>
                  </div>
                );
              } else {
                return (
                  <div>
                    <Stack
                      horizontal
                      tokens={stackTokens}
                      styles={stackStyles}
                    >
                      <Stack {...columnstyle}>
                        <TextField
                          label="User"
                          width="200px"
                          id={index + ""}
                          onChange={(e) => this.userchange.call(this, e)}
                          value={user}
                          name="userName"
                          required
                        ></TextField>
                      </Stack>

                      <Stack {...iconcolumnstyle}>
                        <IconButton
                          iconProps={{ iconName: "Add" }}
                          title="Add User"
                          ariaLabel="Add"
                          onClick={this.newuser.bind(this)}
                        />
                      </Stack>

                      <Stack {...iconcolumnstyle}>
                        <IconButton
                          iconProps={{ iconName: "Cancel" }}
                          title="Remove User"
                          ariaLabel="Cancel"
                          onClick={this.removeuser.bind(this, index)}
                        />
                      </Stack>
                    </Stack>
                  </div>
                );
              }
            })}
          </Stack>

          <div className={styles.margintop}>
            <PrimaryButton onClick={this.processInputProvider}>
              {this.state.formData.Id == 0 ? "Add New Provider" : "Submit"}
            </PrimaryButton>
            <DefaultButton onClick={this.hideDialog.bind(this)} text="Close" />
          </div>

        </Dialog>



        <Dialog
          hidden={this.state.hideDeleteDialog}
          dialogContentProps={dialogContentProps}
        >
          <DialogFooter>
            <PrimaryButton onClick={this.deleteItems.bind(this)} text="Yes" />
            <DefaultButton onClick={this.hideDelete.bind(this)} text="No" />
          </DialogFooter>
        </Dialog>

      </div>
    );

  }
}
