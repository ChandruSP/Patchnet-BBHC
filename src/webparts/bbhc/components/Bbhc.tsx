import * as React from "react";
import styles from "./Bbhc.module.scss";
import { IBbhcProps } from "./IBbhcProps";
import { escape } from "@microsoft/sp-lodash-subset";
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/site-users/web";

import { IItemAddResult } from "@pnp/sp/items";

import { PrimaryButton } from "@fluentui/react";
import { Label } from "office-ui-fabric-react/lib/Label";

import { getId } from "office-ui-fabric-react/lib/Utilities";
import { IStackTokens, Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import * as ReactIcons from '@fluentui/react-icons';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';

import {
  Pivot,
  PivotItem,
  PivotLinkSize,
} from "office-ui-fabric-react/lib/Pivot";

import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
} from "office-ui-fabric-react/lib/Dropdown";

import { ExcelRenderer } from "react-excel-renderer";

var folders: IDropdownOption[] = [];
const currentYear = new Date().getFullYear();
const fileId = getId("anInput");
const options: IDropdownOption[] = [
  {
    key: "fruitsHeader",
    text: "Fruits",
    itemType: DropdownMenuItemType.Header,
  },
  { key: "apple", text: "Apple" },
  { key: "banana", text: "Banana" },
  { key: "orange", text: "Orange", disabled: true },
  { key: "grape", text: "Grape" },
  { key: "divider_1", text: "-", itemType: DropdownMenuItemType.Divider },
  {
    key: "vegetablesHeader",
    text: "Vegetables",
    itemType: DropdownMenuItemType.Header,
  },
  { key: "broccoli", text: "Broccoli" },
  { key: "carrot", text: "Carrot" },
  { key: "lettuce", text: "Lettuce" },
];

import {
  TextField,
  MaskedTextField,
} from "office-ui-fabric-react/lib/TextField";

export interface IBbhcState {
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
    // UsersId: {
    //   results: any[];
    // };
    Users: string;
  };
}

export default class Bbhc extends React.Component<IBbhcProps, IBbhcState> {
  selUsers = [];
  allUsers = [];
  fileObj = null;

  constructor(prop: IBbhcProps, state: IBbhcState) {
    super(prop);
    this.state = {
      providerName: "",
      folders: [],
      subFolders: [],
      cols: [],
      rows: [],
      AllUsers: [''],
      formData: {
        ProviderID: "",
        Title: "",
        LegalName: "",
        // UsersId: {
        //   results: [],
        // },
        Users: ''
      },
    };
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

  componentDidMount() {
    var folderPath = "Shared Documents/2020/Chandru";
    var numbers = [1, 2, 3];
    this.getFolderNames(folderPath);
  }
  getFolderNames = async (folderPath) => {
    var reacthandler = this;
    await sp.web
      .getFolderByServerRelativePath(folderPath)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          reacthandler.processFolderNames(0, data);
        }
        /*for (var k in data) {
          var newfolderPath = folderPath + "/" + data[k].Name;
          if (data[k].ItemCount > 0) {
            folders.push({
              key: data[k].Name,
              text: data[k].Name,
              itemType:
                data[k].ItemCount > 0
                  ? DropdownMenuItemType.Header
                  : DropdownMenuItemType.Normal,
            });
            reacthandler.getFolderNames(newfolderPath);
            folders.push({
              key: "divider_1",
              text: "-",
              itemType: DropdownMenuItemType.Divider,
            });
          } else {
            folders.push({
              key: data[k].Name,
              text: data[k].Name,
              itemType:
                data[k].ItemCount > 0
                  ? DropdownMenuItemType.Header
                  : DropdownMenuItemType.Normal,
            });
          }
        }
        console.log(folders);

        reacthandler.setState({
          folders: folders,
        });*/
      });
  };

  processFolderNames(index, data) {
    var reacthandler = this;

    // reacthandler.createFolder(clonedUrl);
    folders.push({
      key: data[index].Name,
      text: data[index].Name,
      itemType:
        data[index].ItemCount > 0
          ? DropdownMenuItemType.Header
          : DropdownMenuItemType.Normal,
    });
    reacthandler.getFolderNames(data[index].ServerRelativeUrl);
    index = index + 1;
    if (index < data.length) {
      reacthandler.processFolderNames(index, data);
    }

    reacthandler.setState({
      folders: folders,
    });
  }

  // processInputProvider = () => {
  //   var formData = this.state.formData;
  //   if (!formData.ProviderID) {
  //     alert("Provider ID is required");
  //     return;
  //   }
  //   if (!formData.Title) {
  //     alert("Provider name is required");
  //     return;
  //   }
  //   if (!formData.LegalName) {
  //     alert("Legal name is required");
  //     return;
  //   }
  //   if (this.selUsers.length <= 0) {
  //     alert("Select any users");
  //     return;
  //   }
  //   this.getUserData(0);
  // };

  // getUserData(index) {
  //   var that = this;
  //   sp.web.siteUsers
  //     .getByLoginName(this.selUsers[index].id)
  //     .get()
  //     .then((res) => {
  //       this.allUsers.push(res.Id);
  //       index = index + 1;
  //       if (index < that.selUsers.length) {
  //         this.getUserData(index);
  //       } else {
  //         that.addToList();
  //       }
  //     });
  // }



  processInputProvider = () => {
    var formData = this.state.formData;
    if (!formData.ProviderID) {
      alert("Provider ID is required");
      return;
    }
    if (!formData.Title) {
      alert("Provider name is required");
      return;
    }
    if (!formData.LegalName) {
      alert("Legal name is required");
      return;
    }
    for (let index = 0; index < this.state.AllUsers.length; index++) {
      const user = this.state.AllUsers[index];
      if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(user)) {
        formData.Users = formData.Users + user + ';';
      } else {
        alert('User ' + (index + 1) + ' not valid');
        return;
      }
    }
    this.setState({ formData: formData });
    this.addToList();
  };


  addToList() {
    sp.web.lists
      .getByTitle("ProviderDetails")
      .items.add(this.state.formData)
      .then((res) => {
        this.createProvider(this.state.providerName);
      });
  }

  // addToList() {
  //   var formData = this.state.formData;
  //   formData.UsersId.results = this.allUsers;
  //   this.setState({ formData: formData });
  //   sp.web.lists
  //     .getByTitle("ProviderDetails")
  //     .items.add(formData)
  //     .then((res) => {
  //       this.createProvider(this.state.providerName);
  //     });
  // }

  cloneFolder = async () => {
    await this.getFolder("Shared Documents/2020", this.state.providerName);
    alert("Folder Cloned Successfully");
  };

  createProvider = (providerName) => {
    var reacthandler = this;
    sp.web.folders
      .add("Shared Documents/" + currentYear + "/" + providerName)
      .then(function (data) {
        reacthandler.getFolder("TemplateLibrary", providerName);
      });
    alert("Provider is created");
  };

  getFolder = (folderPath, providerName) => {
    var reacthandler = this;
    sp.web
      .getFolderByServerRelativePath(folderPath)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          reacthandler.processFolder(0, data, providerName);
        }
      });
  };

  processFolder(index, data, providerName) {
    var reacthandler = this;
    var clonedUrl = data[index].ServerRelativeUrl.replace(
      "TemplateLibrary",
      "Shared Documents/" + currentYear + "/" + providerName
    );
    // reacthandler.createFolder(clonedUrl);
    sp.web.folders.add(clonedUrl).then((res) => {
      reacthandler.getFolder(data[index].ServerRelativeUrl, providerName);
      index = index + 1;
      if (index < data.length) {
        reacthandler.processFolder(index, data, providerName);
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
    } else {
      reacthandler.fileObj = null;
    }
  };

  uploadToList() {
    var reacthandler = this;
    if (!reacthandler.fileObj) {
      return;
    }
    ExcelRenderer(reacthandler.fileObj, (err, resp) => {
      if (resp && resp.rows) {
        for (let index = 0; index < resp.rows.length; index++) {
          reacthandler.createProvider(resp.rows[index][0]);
        }
      }
    });
  }

  private _getPeoplePickerItems(items: any[]) {
    this.selUsers = items;
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
    allusers.push('');
    this.setState({ AllUsers: allusers });
  }

  public render(): React.ReactElement<IBbhcProps> {

    const stackTokens: IStackTokens = {
      childrenGap: 4,
    };


    const columnstyle: Partial<IStackProps> = {
      tokens: {
        childrenGap: 5
      },
      styles: {
        root: {
          width: 300
        }
      },
    };

    return (
      <div className={styles.bbhc}>
        <Pivot linkSize={PivotLinkSize.large}>
          <PivotItem headerText="Add Provider">

            <Stack {...columnstyle}>

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
                label="Legal Name"
                width="200px"
                onChange={(e) => this.inputChangeHandler.call(this, e)}
                value={this.state.formData.LegalName}
                name="LegalName"
              ></TextField>

              {/* <PeoplePicker
                context={this.props.context}
                titleText="Users"
                personSelectionLimit={100}
                groupName={""}
                showtooltip={true}
                isRequired={false}
                disabled={false}
                selectedItems={this._getPeoplePickerItems.bind(this)}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              /> */}

              {
                this.state.AllUsers.map((user, index) => {
                  if (index == this.state.AllUsers.length - 1) {
                    return <div><TextField
                      label={"User " + (index + 1)}
                      width="200px"
                      id={index + ''}
                      onChange={(e) => this.userchange.call(this, e)}
                      value={user}
                      name="LegalName"
                    ></TextField><a onClick={this.newuser.bind(this)}>Add</a></div>

                  } else {
                    return <div><TextField
                      label={"User " + (index + 1)}
                      width="200px"
                      id={index + ''}
                      onChange={(e) => this.userchange.call(this, e)}
                      value={user}
                      name="LegalName"
                    ></TextField><a onClick={this.removeuser.bind(this, index)}>Remove</a></div>
                  }

                })
              }


            </Stack>

            <div className={styles["margin-top-20"]}>
              <PrimaryButton onClick={this.processInputProvider}>
                Add a New Provider
              </PrimaryButton>
            </div>
          </PivotItem>
          <PivotItem headerText="Upload Excel File">
            <input type="file" id={fileId} onChange={this.uploadFile}></input>

            <Label htmlFor={fileId}>
              <Label>Upload Providers List</Label>
            </Label>

            <PrimaryButton text="Clone" onClick={this.uploadToList.bind(this)} />

          </PivotItem>
        </Pivot>



        {/*<div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>
                Customize SharePoint experiences using Web Parts.
              </p>
              <p className={styles.description}>
                {escape(this.props.description)}
              </p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
    </div>*/}
      </div>
    );
  }
}
