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

import { IItemAddResult } from "@pnp/sp/items";

import { PrimaryButton } from "@fluentui/react";
import { Label } from "office-ui-fabric-react/lib/Label";

import { getId } from "office-ui-fabric-react/lib/Utilities";

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
}

export default class Bbhc extends React.Component<IBbhcProps, IBbhcState> {
  constructor(prop: IBbhcProps, state: IBbhcState) {
    super(prop);
    this.state = {
      providerName: "",
      folders: [],
      subFolders: [],
      cols: [],
      rows: [],
    };
  }
  providerNameChange = (event) => {
    this.setState({ providerName: event.target.value });
  };

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

  processInputProvider = () => {
    this.createProvider(this.state.providerName);
  };

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
    let fileObj = event.target.files[0];
    ExcelRenderer(fileObj, (err, resp) => {
      if (resp && resp.rows) {
        for (let index = 0; index < resp.rows.length; index++) {
          reacthandler.createProvider(resp.rows[index][0]);
        }
      }
    });
  };

  public render(): React.ReactElement<IBbhcProps> {
    return (
      <div className={styles.bbhc}>
        <div>
          <h2>Add Provider</h2>
          <div>
            <TextField
              label="Provider Name"
              onChange={this.providerNameChange}
              width="200px"
            ></TextField>
          </div>
          <div>
            <PrimaryButton onClick={this.processInputProvider}>
              Add a New Provider
            </PrimaryButton>
          </div>
        </div>
        <div>
          <h2>Clone Folder</h2>

          <div>
            <PrimaryButton onClick={this.cloneFolder}>Clone</PrimaryButton>
          </div>
        </div>
        <div>
          <h2>Upload Excel File</h2>
          <input type="file" id={fileId} onChange={this.uploadFile}></input>

          <Label htmlFor={fileId}>
            <Label>Attach File</Label>
          </Label>
        </div>
        <div>
          <h2>Add File</h2>
          <div>
            <Dropdown
              placeholder="Select an option"
              label="Basic uncontrolled example"
              options={this.state.folders}
            />
          </div>
          <div>
            <Dropdown
              placeholder="Select an option"
              label="Basic uncontrolled example"
              options={options}
            />
          </div>
        </div>
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
