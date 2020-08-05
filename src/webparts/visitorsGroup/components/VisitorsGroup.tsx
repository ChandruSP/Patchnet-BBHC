import * as React from "react";
import { escape } from "@microsoft/sp-lodash-subset";

import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups";

import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";

import styles from "./VisitorsGroup.module.scss";


import { IItemAddResult } from "@pnp/sp/items";

import { Label } from "office-ui-fabric-react/lib/Label";

import { Announced } from "office-ui-fabric-react/lib/Announced";
import {
  TextField,
  ITextFieldStyles,
} from "office-ui-fabric-react/lib/TextField";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  IDetailsListStyles,
} from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import {
  CommandBar,
  ICommandBarStyles,
} from "office-ui-fabric-react/lib/CommandBar";

import { ExcelRenderer } from "react-excel-renderer";
import { useId, useBoolean } from "@uifabric/react-hooks";

import {
  IStackTokens,
  Stack,
  IStackProps,
  IStackStyles,
} from "office-ui-fabric-react/lib/Stack";
import * as ReactIcons from "@fluentui/react-icons";
import { mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
import { IconButton } from "@fluentui/react/lib/Button";

import {
  Dialog,
  DialogType,
  DialogFooter,
  IDialogStyles,
} from "office-ui-fabric-react/lib/Dialog";
import {
  DefaultButton,
  PrimaryButton,
} from "office-ui-fabric-react/lib/Button";

import { getId } from "office-ui-fabric-react/lib/Utilities";

import "alertifyjs";

import "../../../ExternalRef/CSS/style.css";
import "../../../ExternalRef/CSS/alertify.min.css";
var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");

import { Image, IImageProps } from "office-ui-fabric-react/lib/Image";
import "@pnp/sp/sputilities";
import { IEmailProperties } from "@pnp/sp/sputilities";

import { IVisitorsGroupProps } from './IVisitorsGroupProps';


export interface IVisitorsGroupState {
  selectionDetails: string;
  items: any[];
  allItems: any[];
  hideDeleteDialog: boolean;
  hideAddDialog: boolean;
  hideSyncDialog: boolean;
  email: string;
  syncUserDetails: string;
  syncUsers: any[];
}

export default class VisitorsGroup extends React.Component<IVisitorsGroupProps, IVisitorsGroupState> {

  visitorsGroupName = 'BBHC Provider SharePoint Viewers';

  visitorsList = 'VisitorsDetails';
  // redirectURL = 'https://bbhcsyncvisitorstolist20200804061631.azurewebsites.net/BBHCVisitors/Index?id=';
  redirectURL = 'http://localhost:51130/BBHCVisitors/Index?id=';

  loginNamePrefix = 'i:0#.f|membership|';
  loginNameSuffix = '#ext#@chandrudemo.onmicrosoft.com';


  private _selection: Selection;
  private _columns: IColumn[];

  constructor(props) {
    super(props);

    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      },
    });


    alertify.set("notifier", "position", "top-right");

    this.state = {
      selectionDetails: '',
      items: [],
      allItems: [],
      hideDeleteDialog: true,
      hideAddDialog: true,
      hideSyncDialog: true,
      email: '',
      syncUserDetails: '',
      syncUsers: []
    };

    this.syncUserDetails();

    this._selection = new Selection({
      onSelectionChanged: () =>
        this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    this._columns = [
      {
        key: "column1",
        name: "User Name",
        fieldName: "Title",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
      },
      {
        key: "column2",
        name: "Email",
        fieldName: "Email",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
      }
    ];
    this.loadVisitors();
  }

  syncUserDetails = () => {
    sp.web.lists
      .getByTitle(this.visitorsList)
      .items.filter("IsSync eq '0' and InvitationAccept eq '1'")
      .get()
      .then((res) => {
        this.setState({ syncUserDetails: res.length + ' user(s) to sync', syncUsers: res });
        if (this.state.syncUsers.length == 0) {
          this.setState({ hideSyncDialog: true });
        }
      });
  }

  loadVisitors() {
    sp.web.siteGroups.getByName(this.visitorsGroupName).users.get().then((result) => {
      var allItems = this.state.items;
      allItems = [];
      for (let index = 0; index < result.length; index++) {
        const element = result[index];
        allItems.push({
          Id: element.Id,
          Title: element.Title,
          Email: element.Email
        });
      }
      this.setState({
        items: allItems,
        allItems: allItems,
        selectionDetails: this._getSelectionDetails(),
      });
    });
  }


  _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const newColumns: IColumn[] = this._columns.slice();
    const currColumn: IColumn = newColumns.filter(
      (currCol) => column.key === currCol.key
    )[0];
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
    const newItems = this._copyAndSort(
      items,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    this.setState({
      items: newItems,
    });
  };


  _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    const key = columnKey as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
      );
  }

  private _onFilter = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ): void => {
    this.setState({
      items: text
        ? this.state.allItems.filter(
          (i) => i.Title.toLowerCase().indexOf(text) > -1
        )
        : this.state.allItems,
    });
  };

  _onAddRow() {

    // this.props.graphClient
    //   .api('/me')
    //   .get()
    //   .then((content: any) => {
    //     debugger;
    //   })
    //   .catch(err => {
    //   });

    this.setState({ hideAddDialog: false, email: '' });
  }

  submitUser = () => {
    if (!this.state.email) {
      alertify.error('Email is required');
      return;
    }
    var formData = {
      Title: this.state.email,
      IsSync: false,
      InvitationAccept: false
    };

    sp.web.lists
      .getByTitle(this.visitorsList)
      .items.add(formData)
      .then((res) => {

        var inviteData = {
          "invitedUserEmailAddress": this.state.email,
          "sendInvitationMessage": true,
          "inviteRedirectUrl": this.redirectURL + btoa(res.data.Id + '-Id')
        };

        this.props.graphClient
          .api('/invitations')
          .post(inviteData)
          .then((content: any) => {
            alertify.success('Invitation sent successfully.');
            this.setState({ email: '', hideAddDialog: true });
          })
          .catch(err => {
            alertify.error('Error while sending invitation.');
          });
      });
  }

  hideDelete() {
    this.setState({ hideDeleteDialog: true });
  }

  _onDeleteRow() {
    this.setState({ hideDeleteDialog: false });
  }

  deleteItems = () => {
    var selItems = this._selection.getSelection();
    if (selItems.length > 0) {
      this.deleteUser(0, selItems);
    } else {
      this.setState({ hideDeleteDialog: true });
    }
  }

  deleteUser = (index, item) => {
    sp.web.siteGroups.getByName(this.visitorsGroupName).users.removeById(item[index].Id).then((res) => {
      index = index + 1;
      if (index < item.length) {
        this.deleteUser(index, item);
      } else {
        this.loadVisitors();
        this.setState({ hideDeleteDialog: true });
        alertify.success('Users(s) deleted successfully');
      }
    });
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();
    switch (selectionCount) {
      case 0:
        return "No items selected";
      case 1:
        return "1 item selected";
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onItemInvoked = (item: IVisitorsGroupProps): void => {
    alert(`Item invoked: ${item.Title}`);
  };

  inputChangeHandler(e) {
    this.setState({
      email: e.target.value
    });
  }


  _syncUserPopup() {
    if (this.state.syncUsers.length > 0) {
      this.setState({ hideSyncDialog: false });
    }
  }

  closesyncpopup = () => {
    this.setState({ hideSyncDialog: true });
  }

  syncsingleuser = (index) => {
    var user = this.state.syncUsers[index];
    var emailId = this.loginNamePrefix + user.Title.replace('@', '_') + this.loginNameSuffix;
    sp.web.ensureUser(emailId).then((result) => {
      if (result) {
        sp.web.siteGroups.getByName(this.visitorsGroupName).users
          .add(result.data.LoginName).then((d) => {
            user.IsSync = true;
            sp.web.lists
              .getByTitle(this.visitorsList)
              .items.getById(user.Id)
              .update(user)
              .then((res) => {
                alertify.success('User synced successfully.');
                this.loadVisitors();
                this.syncUserDetails();
              });
          });
      } else {
        alertify.error('User cannot be found.');
      }
    }).catch((err) => {
      alertify.error('User cannot be found.');
    });
  }

  syncalluser = () => {
    if (this.state.syncUsers.length > 0) {
      this.synconebyone(0);
    }
  }

  synconebyone = (index) => {
    var user = this.state.syncUsers[index];
    var emailId = this.loginNamePrefix + user.Title.replace('@', '_') + this.loginNameSuffix;
    sp.web.ensureUser(emailId).then((result) => {
      if (result) {
        sp.web.siteGroups.getByName(this.visitorsGroupName).users
          .add(result.data.LoginName).then((d) => {
            user.IsSync = true;
            sp.web.lists
              .getByTitle(this.visitorsList)
              .items.getById(user.Id)
              .update(user)
              .then((res) => {
                index = index + 1;
                if (index < this.state.syncUsers.length) {
                  this.synconebyone(index);
                } else {
                  alertify.success('User(s) synced successfully.');
                  this.loadVisitors();
                  this.syncUserDetails();
                }
              });
          });
      }
    }).catch((err) => {
      alertify.error('User cannot be found.');
    });
  }

  updateToList = () => {
    for (let index = 0; index < this.state.syncUsers.length; index++) {
      const user = this.state.syncUsers[index];
      sp.web.lists
        .getByTitle("ProviderDetails")
        .items.getById(user.Id)
        .update(user)
        .then((res) => {
        });
    }
  }

  public render(): React.ReactElement<IVisitorsGroupProps> {

    const exampleChildClass = mergeStyles({
      display: "block",
      marginBottom: "0",
    });

    const commandBarStyles: Partial<ICommandBarStyles> = {
      root: {
        marginBottom: 15,
        padding: 0,
        selectors: {
          ".ms-Button": {
            borderWidth: 0,
            marginRight: 5,
            marginLeft: 5,
            fontFamily: "Poppins, sans-serif",
          },
        },
      },
    };

    const textFieldStyles: Partial<ITextFieldStyles> = {
      root: { maxWidth: "100%", fontFamily: "Poppins, sans-serif" },
    };

    const _renderItemColumn = (item, index: number, column: IColumn) => {
      const fieldContent = item[column.fieldName] as string;
      return <span>{fieldContent}</span>;
    };

    const gridStyles: Partial<IDetailsListStyles> = {
      headerWrapper: [
        {
          selectors: {
            ".ms-DetailsHeader-cell": {
              backgroundColor: "rgb(243, 242, 241)",
              fontFamily: "Poppins, sans-serif",
              marginTop: 3,
            },
            ".ms-DetailsHeader-cellName": {
              fontWeight: 500,
              fontSize: "14px",
            },
          },
        },
      ],
      contentWrapper: [
        {
          selectors: {
            ".ms-DetailsRow": {
              fontFamily: "Poppins, sans-serif",
            },
          },
        },
      ],
    };

    const dialogStyles: Partial<IDialogStyles> = {
      main: [
        {
          fontFamily: "Poppins, sans-serif",
          selectors: {
            ".ms-Dialog-title": {
              fontFamily: "Poppins, sans-serif",
            },
            ".ms-Dialog-subText": {
              fontFamily: "Poppins, sans-serif",
            },
          },
        },
      ],
    };

    const dialogContentProps = {
      type: DialogType.normal,
      title: "Delete",
      closeButtonAriaLabel: "Close",
      subText: "Do you want to delete the selected user(s)?",
    };

    const modelProps = {
      isBlocking: true,
      topOffsetFixed: false,
    };


    const iconcolumnstyle: Partial<IStackProps> = {
      tokens: {
        childrenGap: 5,
      },
      styles: {
        root: {
          width: 30,
        },
      },
    };

    const stackTokens: IStackTokens = {
      childrenGap: 4,
    };
    const stackStyles: Partial<IStackStyles> = {
      root: {
        // width: 600,
      },
    };

    const columnstyle: Partial<IStackProps> = {
      tokens: {
        childrenGap: 5,
      },
      styles: {
        root: {
          width: "100%",
          // paddingTop: 10,
        },
      },
    };

    return (
      <div>
        <CommandBar
          styles={commandBarStyles}
          items={[
            {
              key: "addRow",
              text: "Add",
              iconProps: { iconName: "Add" },
              onClick: this._onAddRow.bind(this),
            },
            // {
            //   key: "deleteRow",
            //   text: "Delete user(s)",
            //   iconProps: { iconName: "Delete" },
            //   onClick: this._onDeleteRow.bind(this),
            // },
            // {
            //   key: 'sync',
            //   text: this.state.syncUserDetails,
            //   iconProps: { iconName: 'Upload', className: this.state.syncUsers.length == 0 ? '' : styles.sync_background },
            //   onClick: this._syncUserPopup.bind(this),
            //   className: this.state.syncUsers.length == 0 ? '' : styles.sync_background
            // },
          ]}
        />

        <Fabric>
          <div className={styles.announcement}>
            <div className={exampleChildClass}>
              {this.state.selectionDetails}
            </div>
            <Announced message={this.state.selectionDetails} />

            <TextField
              prefix="Filter by user name:"
              onChange={this._onFilter.bind(this)}
              styles={textFieldStyles}
              className={styles.searchTextbox}
            />
            <Announced
              message={`Number of items after filter applied: ${this.state.items.length}.`}
            />
          </div>
          <div className={styles.tableContainer}>
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
                styles={gridStyles}
              />
            </MarqueeSelection>
          </div>
        </Fabric>

        <Dialog
          hidden={this.state.hideDeleteDialog}
          dialogContentProps={dialogContentProps}
          styles={dialogStyles}
        >
          <DialogFooter>
            <PrimaryButton
              onClick={this.deleteItems.bind(this)}
              text="Yes"
              className={styles.button_primary}
            />
            <DefaultButton onClick={this.hideDelete.bind(this)} text="No" />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={this.state.hideAddDialog}
          modalProps={modelProps}
          minWidth="400px"
          styles={dialogStyles}
        >
          <Stack {...columnstyle}>

            <TextField
              label="Email ID"
              onChange={(e) => this.inputChangeHandler.call(this, e)}
              name="email"
              value={this.state.email}
              required
              className={styles.input_field}
            ></TextField>

          </Stack>

          <DialogFooter>
            <PrimaryButton
              onClick={this.submitUser}
              className={styles.button_primary}
            >Submit</PrimaryButton>
            <DefaultButton onClick={(e) => this.setState({ hideAddDialog: true })} text="Close" />
          </DialogFooter>

        </Dialog>




        <Dialog
          hidden={this.state.hideSyncDialog}
          modalProps={modelProps}
          minWidth="400px"
          styles={dialogStyles}
        >

          {
            this.state.syncUsers.map((user, index) => {
              return (
                <div>
                  <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                    <Stack {...columnstyle}>
                      <TextField
                        value={user.Title}
                        readOnly={true}
                        className={styles.input_field}
                      ></TextField>
                    </Stack>

                    <Stack {...iconcolumnstyle}>
                      <IconButton
                        iconProps={{ iconName: "Upload" }}
                        title="Sync User"
                        ariaLabel="Sync"
                        onClick={this.syncsingleuser.bind(this, index)}
                        className={styles.primary_button}
                      />
                    </Stack>
                  </Stack>
                </div>
              )
            })
          }

          <DialogFooter>
            <PrimaryButton
              onClick={this.syncalluser}
              className={styles.button_primary}
            >Sync All</PrimaryButton>
            <DefaultButton onClick={this.closesyncpopup.bind(this)} text="Close" />
          </DialogFooter>

        </Dialog>

      </div>
    );
  }
}
