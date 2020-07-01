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
  fileName: "";
}

export default class ProviderDocuments extends React.Component<
  IProviderDocumentsProps,
  IBbhcState
> {
  currentYear = new Date().getFullYear();
  rootFolder = "Providers Library";
  userName = "";

  constructor(props) {
    super(props);

    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      },
    });

    alertify.set("notifier", "position", "top-right");
    this.state = {
      folders: [
        {
          key: "Reports - Deaf and Hard of Hearing",
          text: "Reports - Deaf and Hard of Hearing",
        },
        {
          key: "Reports - Voter Registration",
          text: "Reports - Voter Registration",
        },
        {
          key: "Reports - Miscellaneous",
          text: "Reports - Miscellaneous",
        },
        {
          key: "Reports - Insurance and Licensing",
          text: "Reports - Insurance and Licensing",
        },
        {
          key: "Incidents and Complaints - Submissions",
          text: "Incidents and Complaints - Submissions",
        },
        {
          key: "Incidents and Complaints - Corrective Action Plans",
          text: "Incidents and Complaints - Corrective Action Plans",
        },
        {
          key: "Customer Satisfaction Surveys",
          text: "Customer Satisfaction Surveys",
        },
        {
          key: "Contract Monitoring - Submission for Desk Review",
          text: "Contract Monitoring - Submission for Desk Review",
        },
        {
          key: "Contract Monitoring - Corrective Action Plans and Follow-up",
          text: "Contract Monitoring - Corrective Action Plans and Follow-up",
        },
        {
          key: "Billing - Invoice",
          text: "Billing - Invoice",
        },
        {
          key: "Billing - Invoice Supportive Documentation",
          text: "Billing - Invoice Supportive Documentation",
        },
        {
          key: "Billing - Yearly Audits",
          text: "Billing - Yearly Audits",
        },
      ],
      destinationPath: [],
      file: null,
      selectedPath: "",
      fileName: "",
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
          // that.getFolders(res[0].Title);
        } else {
          that.setState({ folders: [] });
        }
      });
  }

  // getFolders(folderName) {
  //   var url = this.sharedDocument + "/" + currentYear + "/" + folderName;
  //   var folders = [];
  //   var that = this;
  //   var allFolders = that.state.folders;
  //   sp.web
  //     .getFolderByServerRelativePath(url)
  //     .folders.get()
  //     .then(function (data) {
  //       if (data.length > 0) {
  //         for (let index = 0; index < data.length; index++) {
  //           folders.push({
  //             key: folderName,
  //             text: data[index].Name,
  //           });
  //         }
  //         allFolders.splice(allFolders.length, 0, folders);
  //         that.setState({ folders: allFolders });
  //       }
  //     });
  // }

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
        var folderName =
          "FY " + (this.currentYear - 1) + "-" + this.currentYear;

        var folderPath =
          this.rootFolder + "/" + folderName + "/" + this.userName + "/";
        var destinationPaths = selectedPath.split(" - ");
        for (let index = 0; index < destinationPaths.length; index++) {
          folderPath = folderPath + destinationPaths[index] + "/";
        }
        var that = this;
        sp.web
          .getFolderByServerRelativeUrl(folderPath)
          .files.add(that.state.file.name, that.state.file, true)
          .then(function (result) {
            alertify.success("File uploaded successfully");
          });
      } else {
        alertify.error("Select any folder");
      }
    } else {
      alertify.error("Select any file");
    }
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
      this.setState({ selectedPath: item.text });
    };

    return (
      <div>
        <h2>Add File</h2>
        <div>
          {
            // this.state.folders.map((folder, index) => {
            //   return <Dropdown
            //     placeholder="Select an option"
            //     label="Folders"
            //     options={folder}
            //     onChange={dropdownChange}
            //     id={index + ""}
            //   />
            // })

            <Dropdown
              placeholder="Select an option"
              label="Submission Types"
              options={this.state.folders}
              onChange={dropdownChange}
              style={{ width: "500px" }}
            />
          }
        </div>
        [ style={{ visibility: "hidden" }}]
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
