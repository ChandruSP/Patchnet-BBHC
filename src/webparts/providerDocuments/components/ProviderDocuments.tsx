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
import { Link } from 'office-ui-fabric-react/lib/Link';

import { getId } from "office-ui-fabric-react/lib/Utilities";

import "alertifyjs";

import "../../../ExternalRef/CSS/style.css";
import "../../../ExternalRef/CSS/alertify.min.css";
var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");

import { Image, IImageProps } from "office-ui-fabric-react/lib/Image";
import "@pnp/sp/sputilities";
import { IEmailProperties } from "@pnp/sp/sputilities";


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
  selectedProvider: string;
  notes: string;
  fileName: "";
  previousyeardata: any[];
  allProviders: any[];
}

export default class ProviderDocuments extends React.Component<
  IProviderDocumentsProps,
  IBbhcState
  > {
  currentYear = new Date().getFullYear();
  rootFolder = "Providers Library";
  templateLibrary = "TemplateLibrary";
  generalSubmission = 'general submission';
  generalSubmissionChanged = false;


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
      selectedProvider: "",
      fileName: "",
      notes: "",
      previousyeardata: [],
      allProviders: []
    };
    this.getProviderMetaData();
  }

  getProviderMetaData() {
    var that = this;
    sp.web.lists
      .getByTitle("ProviderDetails")
      .items.select("Title", "ContractId", "TemplateType")
      .filter(
        "substringof('" +
        this.props.currentContext.pageContext.user.email.toLowerCase() +
        "',Users)"
      )
      .get()
      .then((res) => {
        if (res.length > 0) {
          var currentMonth = new Date().getMonth() + 1;
          var stryear = that.currentYear;
          if (currentMonth < 7) {
            stryear = that.currentYear - 1;
          }
          var previousyeardata = that.state.previousyeardata;
          var allProviders = that.state.allProviders;

          var dataLoaded = false;
          for (let j = 0; j < res.length; j++) {
            const providerData = res[j];
            var contract = providerData.ContractId.substr(providerData.ContractId.length - 2, 2);
            if (contract != stryear.toString().substr(2, 2)) {
              var nextyear = parseInt(contract) + 1;
              var currentyearprefix = that.currentYear.toString().substr(0, 2);
              previousyeardata.push({
                Title: "FY " + (currentyearprefix + contract) + "-" + (currentyearprefix + nextyear),
                URL: that.props.siteUrl + "/" + this.rootFolder + "/FY " + (currentyearprefix + contract) + "-" + (currentyearprefix + nextyear) + "/" + providerData.Title
              });
            } else {
              if (!dataLoaded) {
                dataLoaded = true;
                that.loadUploadFolders(providerData.TemplateType);
              }
              allProviders.push({
                key: providerData.Title,
                text: providerData.Title
              });
            }
          }
          that.setState({ previousyeardata: previousyeardata, allProviders: allProviders });
        } else {
          that.setState({ folders: [] });
        }
      });
  }

  loadUploadFolders(templateType) {
    var that = this;
    sp.web.lists
      .getByTitle("UploadFolders")
      .items.select("Title", "TemplateType")
      .filter("TemplateType eq '" + templateType + "'")
      .get()
      .then((res) => {
        var allFolders = that.state.folders;
        allFolders = [];
        for (let index = 0; index < res.length; index++) {
          var cleartext = res[index].Title.replace(' - Upload', '');
          var url = cleartext.replace(' - ', '/');
          allFolders.push({
            key: url,
            text: cleartext,
          });
        }
        that.setState({ folders: allFolders });
      });
  }

  // getFolders(folderName, templateType, displayName) {
  //   var url = this.templateLibrary + "/" + templateType;
  //   if (folderName) {
  //     url = url + "/" + folderName;
  //   }
  //   var that = this;
  //   var allFolders = that.state.folders;
  //   sp.web
  //     .getFolderByServerRelativePath(url)
  //     .folders.get()
  //     .then(function (data) {
  //       if (data.length > 0) {
  //         for (let index = 0; index < data.length; index++) {
  //           var text = '';
  //           var cleartext = data[index].Name.replace(' - Upload', '')
  //           if (displayName) {
  //             text = displayName + ' - ' + cleartext;
  //           } else {
  //             text = cleartext;
  //           }
  //           if (data[index].Name.toLocaleLowerCase().indexOf('upload') > 0) {
  //             allFolders.push({
  //               key: folderName + '/' + cleartext,
  //               text: text,
  //             });
  //           }
  //           that.setState({ folders: allFolders });
  //           that.getFolders(folderName + '/' + data[index].Name, templateType, text);
  //         }
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
      if (!this.state.selectedProvider) {
        alertify.error("Select any provider");
        return;
      }
      var selectedPath = this.state.selectedPath;
      if (selectedPath) {
        if (selectedPath.toLocaleLowerCase().indexOf(this.generalSubmission) > 0) {
          if (!this.state.notes) {
            alertify.error('Notes is required');
            return;
          }
        }
        var currentMonth = new Date().getMonth() + 1;
        var stryear = this.currentYear + "-" + (this.currentYear + 1);
        if (currentMonth < 7) {
          stryear = (this.currentYear - 1) + "-" + this.currentYear;
        }
        var folderName = "FY " + stryear + "/" + this.state.selectedProvider;

        var folderPath = this.rootFolder + "/" + folderName + "/" + selectedPath;
        var that = this;
        sp.web
          .getFolderByServerRelativeUrl(folderPath)
          .files.add(that.state.file.name, that.state.file, true)
          .then(function (result) {
            if (selectedPath.toLocaleLowerCase().indexOf(that.generalSubmission) > 0) {
              result.file.listItemAllFields.get().then(function (fileData) {
                sp.web.lists.getByTitle(that.rootFolder).items.getById(fileData.Id).update({ FileNotes: that.state.notes }).then(function () {

                  sp.web.lists
                    .getByTitle("EmailConfig")
                    .items
                    .get()
                    .then((res) => {
                      var filepath = that.props.currentContext.pageContext.web.absoluteUrl + '/' + folderPath + that.state.file.name;
                      var to = res[0].To.split(';');
                      var cc = [];
                      if (res[0].CC) {
                        cc = res[0].CC.split(';');
                      }
                      var bcc = [];
                      if (res[0].BCC) {
                        bcc = res[0].BCC.split(';');
                      }
                      const emailProps: IEmailProperties = {
                        To: to,
                        CC: cc,
                        BCC: bcc,
                        Subject: res[0].Subject,
                        Body: "New file is uploaded in the general submission folder for the <a href='" + filepath + "'>" + that.state.selectedProvider + "</a> provider.\n\nNotes : " + that.state.notes,
                        AdditionalHeaders: {
                          "content-type": "text/html"
                        }
                      };
                      sp.utility.sendEmail(emailProps);
                    });

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
      if (!this.generalSubmissionChanged) {
        this.generalSubmissionChanged = true;
        var folders = this.state.folders;
        var gindex = -1;
        for (let index = 0; index < folders.length; index++) {
          const folder = folders[index];
          if (folder.text == "General Submissions") {
            gindex = index;
            break;
          }
        }
        if (gindex >= 0) {
          var data = folders[gindex];
          folders.splice(gindex, 1);
          folders.splice(folders.length, 0, data);
          this.setState({ folders: folders });
        }
      }
      this.setState({ selectedPath: item.key.toString() });
    };

    const providerChange = (
      event: React.FormEvent<HTMLDivElement>,
      item: IDropdownOption
    ): void => {
      this.setState({ selectedProvider: item.key.toString() });
    };


    return (
      <div>
        <h2>Add File</h2>
        <div>

          {
            <Dropdown
              placeholder="Select an provider"
              label="Providers"
              options={this.state.allProviders}
              onChange={providerChange}
              style={{ width: "700px" }}
            />
          }

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

        <div>
          {
            this.state.previousyeardata.map((provider) => {
              return <div><Link href={provider.URL} target="_blank">{provider.Title}</Link><br></br></div>
            })
          }
        </div>

        <PrimaryButton text="Upload" onClick={this.uploadFile.bind(this)} />
      </div>
    );
  }
}
