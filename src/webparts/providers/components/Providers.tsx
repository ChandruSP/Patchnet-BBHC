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

  providerName: "";
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
    text: "Contract Providers",
  }, {
    key: "Agreement Providers",
    text: "Agreement Providers",
  }];

  contributePermission = null;
  readPermission = null;
  currentUser = null;
  userIds = [];

  constructor(props) {
    super(props);
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
      { key: 'column1', name: 'Provider Name', fieldName: 'Title', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'Legal Name', fieldName: 'LegalName', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column3', name: 'Provider ID', fieldName: 'ProviderID', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column4', name: 'Template Type', fieldName: 'TemplateType', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column5', name: 'Contract Id', fieldName: 'ContractId', minWidth: 100, maxWidth: 200, isResizable: true },
    ];

    var that = this;
    that.state = {
      items: [],
      allItems: [],
      selectionDetails: '',
      hideDialog: true,
      hideDeleteDialog: true,

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

    sp.web.lists
      .getByTitle("ProviderDetails")
      .items.select("Title", "LegalName", "ProviderID", "TemplateType", "ContractId", "Id", "Users", "IsDeleted", "Logs").get().then(function (data) {
        var allItems = that.state.items;
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
    this.setState({ AllUsers: allUsers, formData: formData, hideDialog: false });
  }

  hideDialog() {
    location.reload();
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
          that.userIds.push(data.Id);
        });

      } else {
        alertify.error("User " + (index + 1) + " not valid");
        return;
      }
    }
    this.setState({ formData: formData });
    this.addToList(currentYear, this.state.formData);
  };

  addToList(year, formData) {
    if (formData.Id > 0) {
      formData.Logs = formData.Logs + "\n\nUpdated on : " + new Date() + "\nUpdated by : " + this.props.currentContext.pageContext.user.displayName;
      sp.web.lists
        .getByTitle("ProviderDetails")
        .items.getById(formData.Id).update(formData)
        .then((res) => {
          alertify.success("Provider updated");
        });
    } else {
      var currentMonth = new Date().getMonth();
      formData.ContractId = currentMonth >= 7 ? (formData.ContractId + '-' + currentYear) : (formData.ContractId + '-' + (currentYear - 1));
      formData.Logs = "Added on : " + new Date() + "\nAdded by : " + this.props.currentContext.pageContext.user.displayName;
      sp.web.lists
        .getByTitle("ProviderDetails")
        .items.add(formData)
        .then((res) => {
          this.createProvider(formData.Title, year, formData);
        });
    }
  }


  createProvider = (providerName, year, formData) => {
    var reacthandler = this;
    var folderName =
      reacthandler.rootFolder + "/" + "FY " + (year - 1) + "-" + year;
    sp.web.folders.add(folderName + "/" + providerName).then(function (data) {
      reacthandler.getFolder("TemplateLibrary/" + formData.TemplateType, providerName, year, formData);
    });
    alertify.success("Provider is created");
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
    var folderName =
      reacthandler.rootFolder + "/" + "FY " + (year - 1) + "-" + year;
    var clonedUrl = data[index].ServerRelativeUrl.replace(
      "TemplateLibrary/" + formData.TemplateType,
      folderName + "/" + providerName
    );
    // reacthandler.createFolder(clonedUrl);
    sp.web.folders.add(clonedUrl).then((res) => {
      var url = clonedUrl.replace(this.props.currentContext.pageContext.web.serverRelativeUrl + '/', '');
      const spHttpClient: SPHttpClient = this.props.currentContext.spHttpClient;
      var queryUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativeUrl(" + "'" + url + "'" + ")/ListItemAllFields/breakroleinheritance(false)";
      const spOpts: ISPHttpClientOptions = {};
      spHttpClient.post(queryUrl, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
        if (response.ok) {
          var permission = reacthandler.readPermission;
          var sdata = clonedUrl.split('/');
          if (sdata[sdata.length - 1].toLocaleLowerCase().indexOf('upload') > 0) {
            permission = reacthandler.contributePermission;
          }
          for (let index = 0; index < reacthandler.userIds.length; index++) {
            const userId = reacthandler.userIds[index];
            var postUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/_api/web/GetFolderByServerRelativeUrl(' + "'" + url + "'" + ')/ListItemAllFields/roleassignments/addroleassignment(principalid=' + userId + ',roledefid=' + permission + ')';
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

  uploadToList() {
    var reacthandler = this;
    if (!reacthandler.fileObj) {
      alertify.error("Select any file to upload");
      return;
    }
    ExcelRenderer(reacthandler.fileObj, (err, resp) => {
      if (resp && resp.rows) {
        for (let index = 1; index < resp.rows.length; index++) {
          let rowData = resp.rows[index];
          var formdata = {
            Title: rowData[0],
            ProviderID: rowData[1],
            ContractId: rowData[2],
            LegalName: rowData[3],
            Users: rowData[4],
            TemplateType: rowData[5]
          };
          if (formdata.Title) {
            reacthandler.addToList(currentYear, formdata);
          }
        }
      }
    });
  }


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
    this.setState({ AllUsers: allUsers, formData: formData, hideDialog: false });
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
        index = index + 1;
        if (index < items.length) {
          that.updateDeleteTag(index, items);
        } else {
          alertify.success("Provider deleted successfully");
          setTimeout(() => {
            location.reload();
          }, 1000);
        }
      });
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
          width: 300,
          paddingTop: 28,
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
      subText: 'Do you want to these providers?',
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
            label="Filter by name:"
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


        <Dialog hidden={this.state.hideDialog} modalProps={modelProps}>

          <Stack {...columnstyle}>

            <ChoiceGroup defaultSelectedKey={this.state.formData.TemplateType} options={this.templateTypes} onChange={this.templateChange.bind(this)} label="Template Type" />

            <TextField
              label="Provider ID"
              onChange={(e) => this.inputChangeHandler.call(this, e)}
              width="100px"
              name="ProviderID"
              value={this.state.formData.ProviderID}
            ></TextField>

            <TextField
              label="Provider Name"
              onChange={this.providerNameChange}
              width="100px"
              name="Title"
              value={this.state.formData.Title}
            ></TextField>

            <TextField
              label="Contract Id"
              width="200px"
              onChange={(e) => this.inputChangeHandler.call(this, e)}
              value={this.state.formData.ContractId}
              name="ContractId"
            ></TextField>

            <TextField
              label="Legal Name"
              width="200px"
              onChange={(e) => this.inputChangeHandler.call(this, e)}
              value={this.state.formData.LegalName}
              name="LegalName"
            ></TextField>




            {this.state.AllUsers.map((user, index) => {
              if (index == this.state.AllUsers.length - 1) {
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
                        ></TextField>
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

          <DialogFooter>
            <PrimaryButton onClick={this.processInputProvider}>
              Add a New Provider
              </PrimaryButton>
            <DefaultButton onClick={this.hideDialog.bind(this)} text="Close" />
          </DialogFooter>

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
