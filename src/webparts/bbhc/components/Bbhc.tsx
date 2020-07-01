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
import { Image, IImageProps } from "office-ui-fabric-react/lib/Image";

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

import "alertifyjs";

import "../../../ExternalRef/CSS/style.css";
import "../../../ExternalRef/CSS/alertify.min.css";
var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");

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
const attachImageStyles = {
  image: {
    padding: "0px",
  },
};

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
import { Icon } from "office-ui-fabric-react/lib/Icon";

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
  fileName: "";
}

export default class Bbhc extends React.Component<IBbhcProps, IBbhcState> {
  selUsers = [];
  allUsers = [];
  fileObj = null;
  rootFolder = "Providers Library";

  constructor(prop: IBbhcProps, state: IBbhcState) {
    super(prop);

    alertify.set("notifier", "position", "top-right");

    this.state = {
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
        // UsersId: {
        //   results: [],
        // },
        Users: "",
      },
      fileName: "",
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
  //     alertify.error("Provider ID is required");
  //     return;
  //   }
  //   if (!formData.Title) {
  //     alertify.error("Provider name is required");
  //     return;
  //   }
  //   if (!formData.LegalName) {
  //     alertify.error("Legal name is required");
  //     return;
  //   }
  //   if (this.selUsers.length <= 0) {
  //     alertify.error("Select any users");
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
      alertify.error("Provider ID is required");
      return;
    }
    if (!formData.Title) {
      alertify.error("Provider name is required");
      return;
    }
    if (!formData.LegalName) {
      alertify.error("Legal name is required");
      return;
    }
    for (let index = 0; index < this.state.AllUsers.length; index++) {
      const user = this.state.AllUsers[index];
      if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(user)) {
        formData.Users = formData.Users + user + ";";
      } else {
        alertify.error("User " + (index + 1) + " not valid");
        return;
      }
    }
    this.setState({ formData: formData });
    this.addToList(currentYear);
  };

  addToList(year) {
    sp.web.lists
      .getByTitle("ProviderDetails")
      .items.add(this.state.formData)
      .then((res) => {
        this.createProvider(this.state.providerName, year);
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

  // cloneFolder = async () => {
  //   await this.getFolder("Shared Documents/2020", this.state.providerName);
  //   alertify.success("Folder Cloned Successfully");
  // };

  createProvider = (providerName, year) => {
    var reacthandler = this;
    var folderName =
      reacthandler.rootFolder + "/" + "FY " + (year - 1) + "-" + year;
    sp.web.folders.add(folderName + "/" + providerName).then(function (data) {
      reacthandler.getFolder("TemplateLibrary", providerName, year);
    });
    alertify.success("Provider is created");
  };

  getFolder = (folderPath, providerName, year) => {
    var reacthandler = this;
    sp.web
      .getFolderByServerRelativePath(folderPath)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          reacthandler.processFolder(0, data, providerName, year);
        }
      });
  };

  processFolder(index, data, providerName, year) {
    var reacthandler = this;

    var folderName =
      reacthandler.rootFolder + "/" + "FY " + (year - 1) + "-" + year;
    var clonedUrl = data[index].ServerRelativeUrl.replace(
      "TemplateLibrary",
      folderName + "/" + providerName
    );
    // reacthandler.createFolder(clonedUrl);
    sp.web.folders.add(clonedUrl).then((res) => {
      reacthandler.getFolder(data[index].ServerRelativeUrl, providerName, year);
      index = index + 1;
      if (index < data.length) {
        reacthandler.processFolder(index, data, providerName, year);
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
      alertify.console.error("Select any file to upload");
      return;
    }
    sp.web.folders
      .add("Shared Documents/" + (currentYear + 1))
      .then(function (data) {
        ExcelRenderer(reacthandler.fileObj, (err, resp) => {
          if (resp && resp.rows) {
            for (let index = 0; index < resp.rows.length; index++) {
              reacthandler.createProvider(resp.rows[index][0], currentYear + 1);
            }
          }
        });
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
    allusers.push("");
    this.setState({ AllUsers: allusers });
  }

  public render(): React.ReactElement<IBbhcProps> {
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

            <div className={styles["margin-top-20"]}>
              <PrimaryButton onClick={this.processInputProvider}>
                Add a New Provider
              </PrimaryButton>
            </div>
          </PivotItem>
          <PivotItem headerText="Clone Previous Year">
            <input
              type="file"
              id={fileId}
              style={{ visibility: "hidden" }}
              onChange={this.uploadFile}
            ></input>

            <Label htmlFor={fileId}>
              <Label styles={{ root: { padding: "5px" } }}>Attach File</Label>
              <div style={{ display: "flex" }}>
                <Image
                  styles={{ image: { padding: "5px" } }}
                  src={require("./Attach.png")}
                ></Image>
                <Label>{this.state.fileName}</Label>
              </div>
            </Label>

            <PrimaryButton
              text="Clone"
              onClick={this.uploadToList.bind(this)}
            />
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
