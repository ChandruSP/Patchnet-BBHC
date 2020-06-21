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

import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
} from "office-ui-fabric-react/lib/Dropdown";

const currentYear = new Date().getFullYear();
import { IProviderDocumentsProps } from './IProviderDocumentsProps';

export interface IBbhcState {
  folders: any[];
  destinationPath: any[];
  file: any;
}

export default class ProviderDocuments extends React.Component<IProviderDocumentsProps, IBbhcState> {

  currentYear = new Date().getFullYear();
  sharedDocument = "Shared Documents";
  userName = '';

  constructor(props) {
    super(props);
    this.state = {
      folders: [],
      destinationPath: [],
      file: null,
    };
    this.getProviderMetaData();
  }

  getProviderMetaData() {
    var that = this;
    sp.web.lists
      .getByTitle("ProviderDetails")
      .items.select("Title")
      .filter("Users/EMail eq '" + this.props.currentContext.pageContext.user._email + "'").get().then((res) => {
        if (res.length > 0) {
          that.userName = res[0].Title;
          that.getFolders(res[0].Title);
        }
      });
  }

  getFolders(folderName) {
    var url = this.sharedDocument + '/' + currentYear + '/' + folderName;
    var folders = [];
    var that = this;
    var allFolders = that.state.folders;
    sp.web
      .getFolderByServerRelativePath(url)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          for (let index = 0; index < data.length; index++) {
            folders.push({
              key: folderName,
              text: data[index].Name
            });
          }
          allFolders.splice(allFolders.length, 0, folders);
          that.setState({ folders: allFolders });
        }
      });
  }

  fileUpload(e) {
    var files = e.target.files;
    if (files && files.length > 0) {
      this.setState({ file: files[0] });
    } else {
      this.setState({ file: null });
    }
  }

  uploadFile() {
    if (this.state.file) {
      var destinationPaths = this.state.destinationPath;
      if (destinationPaths.length > 0) {

        if (destinationPaths.length != this.state.folders.length) {
          alert('Fill all dropdown values');
          return;
        }

        var folderPath = this.sharedDocument + '/' + this.currentYear + '/' + this.userName + '/';
        for (let index = 0; index < destinationPaths.length; index++) {
          folderPath = folderPath + destinationPaths[index].value + '/';
        }
        var that = this;
        sp.web.getFolderByServerRelativeUrl(folderPath).files.add(that.state.file.name, that.state.file, true)
          .then(function (result) {
            alert('File uploaded successfully');
          });
      } else {
        alert('Select any folder');
      }
    } else {
      alert('Select any file');
    }
  }

  public render(): React.ReactElement<IProviderDocumentsProps> {

    const dropdownChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
      if (item) {
        var dropDownIndex = parseInt(event.target["id"]);
        var destinationPath = this.state.destinationPath;
        var destinationFound = false;
        for (let d = 0; d < destinationPath.length; d++) {
          if (destinationPath[d].index == dropDownIndex) {
            destinationPath[d].value = item.text;
            destinationFound = true;
            break;
          }
        }
        if (!destinationFound) {
          destinationPath.push({
            index: dropDownIndex,
            value: item.text
          });
        }
        var stateFolder = this.state.folders;
        var removeIndexes = [];
        for (let index = dropDownIndex + 1; index < this.state.folders.length; index++) {
          removeIndexes.push(index);
        }
        for (var i = removeIndexes.length - 1; i >= 0; i--) {
          stateFolder.splice(removeIndexes[i], 1);
          for (let d = 0; d < destinationPath.length; d++) {
            if (destinationPath[d].index == removeIndexes[i]) {
              destinationPath.splice(removeIndexes[i], 1);
            }
          }
        }
        this.setState({ folders: stateFolder, destinationPath: destinationPath });
        this.getFolders(item.key + '/' + item.text);
      }
    };

    return (
      <div>
        <h2>Add File</h2>
        <div>
          {
            this.state.folders.map((folder, index) => {
              return <Dropdown
                placeholder="Select an option"
                label="Folders"
                options={folder}
                onChange={dropdownChange}
                id={index + ""}
              />
            })
          }
        </div>

        <input type="file" name="UploadedFile" onChange={(e) => this.fileUpload.call(this, e)} />

        <PrimaryButton text="Post" onClick={this.uploadFile.bind(this)} />


      </div>
    );
  }
}
