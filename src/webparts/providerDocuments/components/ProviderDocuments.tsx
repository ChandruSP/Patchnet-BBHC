import * as React from "react";
import { escape } from "@microsoft/sp-lodash-subset";
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
import "@pnp/sp/files";

import { IItemAddResult } from "@pnp/sp/items";

import { PrimaryButton } from "@fluentui/react";
import { Label } from "office-ui-fabric-react/lib/Label";

import { getId } from "office-ui-fabric-react/lib/Utilities";

import "alertifyjs";

import "../../../ExternalRef/CSS/style.css";
import "../../../ExternalRef/CSS/alertify.min.css";
var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");

import { Image, IImageProps } from "office-ui-fabric-react/lib/Image";


import {
  TextField,
  MaskedTextField,
} from "office-ui-fabric-react/lib/TextField";

import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
} from "office-ui-fabric-react/lib/Dropdown";

const currentYear = new Date().getFullYear();

const fileId = getId("anInput");
import { IProviderDocumentsProps } from "./IProviderDocumentsProps";

export interface IBbhcState {
  folders: any[];
  destinationPath: any[];
  file: any;
  selectedPath: string;
  notes: string;
  fileName: "";
}

export default class ProviderDocuments extends React.Component<
  IProviderDocumentsProps,
  IBbhcState
  > {
  currentYear = new Date().getFullYear();
  rootFolder = "Providers Library";
  userName = "";
  generalSubmission = 'general submission';

  constructor(props) {
    super(props);

    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      },
    });

    alertify.set("notifier", "position", "top-right");
    this.state = {
      folders: [],
      destinationPath: [],
      file: null,
      selectedPath: "",
      fileName: "",
      notes: ""
    };
    this.getProviderMetaData();
  }

  getProviderMetaData() {
    var that = this;
    sp.web.lists
      .getByTitle("ProviderDetails")
      .items.select("Title")
      .filter(
        "substringof('" +
        this.props.currentContext.pageContext.user.email.toLowerCase() +
        "',Users)"
      )
      .get()
      .then((res) => {
        if (res.length > 0) {
          that.userName = res[0].Title;
          that.getFolders(res[0].Title, '');
        } else {
          that.setState({ folders: [] });
        }
      });
  }

  getFolders(folderName, displayName) {
    var url = this.rootFolder + "/" + "FY " + (currentYear - 1) + "-" + currentYear + "/" + folderName;
    var that = this;
    var allFolders = that.state.folders;
    sp.web
      .getFolderByServerRelativePath(url)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          for (let index = 0; index < data.length; index++) {
            var text = '';
            if (displayName) {
              text = displayName + ' - ' + data[index].Name.replace(' - Upload', '');
            } else {
              text = data[index].Name.replace(' - Upload', '');
            }
            if (data[index].Name.toLocaleLowerCase().indexOf('upload') > 0) {
              allFolders.push({
                key: folderName + '/' + data[index].Name,
                text: text,
              });
            }
            that.setState({ folders: allFolders });
            that.getFolders(folderName + '/' + data[index].Name, text);
          }
        }
      });
  }

  fileUpload(e) {
    var files = e.target.files;
    if (files && files.length > 0) {
      this.setState({ file: files[0], fileName: files[0].name });
    } else {
      this.setState({ file: null });
    }
  }

  // uploadFile() {
  //   if (this.state.file) {
  //     var destinationPaths = this.state.destinationPath;
  //     if (destinationPaths.length > 0) {

  //       if (destinationPaths.length != this.state.folders.length) {
  //         alertify.error('Fill all dropdown values');
  //         return;
  //       }

  //       var folderPath = this.sharedDocument + '/' + this.currentYear + '/' + this.userName + '/';
  //       for (let index = 0; index < destinationPaths.length; index++) {
  //         folderPath = folderPath + destinationPaths[index].value + '/';
  //       }
  //       var that = this;
  //       sp.web.getFolderByServerRelativeUrl(folderPath).files.add(that.state.file.name, that.state.file, true)
  //         .then(function (result) {
  //           alertify.success('File uploaded successfully');
  //         });
  //     } else {
  //       alertify.error('Select any folder');
  //     }
  //   } else {
  //     alertify.error('Select any file');
  //   }
  // }

  uploadFile() {
    if (this.state.file) {
      var selectedPath = this.state.selectedPath;
      if (selectedPath) {
        if (selectedPath.toLocaleLowerCase().indexOf(this.generalSubmission) > 0) {
          if (!this.state.notes) {
            alertify.error('Notes is required');
            return;
          }
        }
        var folderName = "FY " + (this.currentYear - 1) + "-" + this.currentYear;
        var folderPath = this.rootFolder + "/" + folderName + "/" + selectedPath;
        var that = this;
        sp.web
          .getFolderByServerRelativeUrl(folderPath)
          .files.add(that.state.file.name, that.state.file, true)
          .then(function (result) {
            if (selectedPath.toLocaleLowerCase().indexOf(that.generalSubmission) > 0) {
              result.file.listItemAllFields.get().then(function (fileData) {
                sp.web.lists.getByTitle(that.rootFolder).items.getById(fileData.Id).update({ FileNotes: that.state.notes }).then(function () {
                  alertify.success("File uploaded successfully");
                })
              });
            } else {
              alertify.success("File uploaded successfully");
            }
          });
      } else {
        alertify.error("Select any folder");
      }
    } else {
      alertify.error("Select any file");
    }
  }

  inputChangeHandler(e) {
    this.setState({
      notes: e.target.value
    });
  }

  public render(): React.ReactElement<IProviderDocumentsProps> {
    // const dropdownChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    //   if (item) {
    //     var dropDownIndex = parseInt(event.target["id"]);
    //     var destinationPath = this.state.destinationPath;
    //     var destinationFound = false;
    //     for (let d = 0; d < destinationPath.length; d++) {
    //       if (destinationPath[d].index == dropDownIndex) {
    //         destinationPath[d].value = item.text;
    //         destinationFound = true;
    //         break;
    //       }
    //     }
    //     if (!destinationFound) {
    //       destinationPath.push({
    //         index: dropDownIndex,
    //         value: item.text
    //       });
    //     }
    //     var stateFolder = this.state.folders;
    //     var removeIndexes = [];
    //     for (let index = dropDownIndex + 1; index < this.state.folders.length; index++) {
    //       removeIndexes.push(index);
    //     }
    //     for (var i = removeIndexes.length - 1; i >= 0; i--) {
    //       stateFolder.splice(removeIndexes[i], 1);
    //       for (let d = 0; d < destinationPath.length; d++) {
    //         if (destinationPath[d].index == removeIndexes[i]) {
    //           destinationPath.splice(removeIndexes[i], 1);
    //         }
    //       }
    //     }
    //     this.setState({ folders: stateFolder, destinationPath: destinationPath });
    //     this.getFolders(item.key + '/' + item.text);
    //   }
    // };

    const dropdownChange = (
      event: React.FormEvent<HTMLDivElement>,
      item: IDropdownOption
    ): void => {
      this.setState({ selectedPath: item.key.toString() });
    };

    return (
      <div>
        <h2>Add File</h2>
        <div>
          {
            <Dropdown
              placeholder="Select an option"
              label="Submission Types"
              options={this.state.folders}
              onChange={dropdownChange}
              style={{ width: "700px" }}
            />

          }

          {
            this.state.selectedPath.toLocaleLowerCase().indexOf(this.generalSubmission) > 0 ?
              <TextField
                label="Notes"
                width="100px"
                onChange={(e) => this.inputChangeHandler.call(this, e)}
                value={this.state.notes}
                name="notes"
              ></TextField>
              : ""
          }

        </div>

        <input
          type="file"
          name="UploadedFile"
          id={fileId}
          onChange={(e) => this.fileUpload.call(this, e)}
          style={{ visibility: "hidden" }}
        />
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
        <PrimaryButton text="Upload" onClick={this.uploadFile.bind(this)} />
      </div>
    );
  }
}
