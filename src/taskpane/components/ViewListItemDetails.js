import * as React from "react";
import { Announced } from "office-ui-fabric-react/lib/Announced";
import {
  TextField,
  PrimaryButton,
  Dropdown,
  DatePicker,
  ActionButton,
  Link,
  TooltipHost,
  TooltipDelay,
  DirectionalHint,
  Label,
  Dialog,
  DialogFooter,
  DialogType,
} from "office-ui-fabric-react";
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import { Text, Spinner } from "office-ui-fabric-react";
import { IconButton, IButtonStyles, DefaultButton } from "office-ui-fabric-react/lib/Button";
import { Link as RouterLink } from "react-router-dom";
import { IOverflowSetItemProps, OverflowSet } from "office-ui-fabric-react/lib/OverflowSet";
import { MyRegistryDetailList } from "./MyRegistryDetailList";
import axios from "axios";
import { StaticConst } from "../helper/Const";
import { AsyncHelper } from "../helper/AsyncHelper";
import moment from "moment";
export class ViewListItemDetails extends React.Component {
  helper = new AsyncHelper(this.props.Authorization);

  constructor(props) {
    super(props);
    this.state = this.initState();
    this.onSaveClick = this.onSaveClick.bind(this);
    //this.onTextchange = this.onTextchange.bind(this);
  }

  initState() {
    return {
      fields: null,
      isdataLoading: false,
      Project_x0020_Name: null,
      Bid_x0020__x0023_: null, //Bid #
      Estimated_x0020_Project_x0020_Va: null,
      Bid_x0020_Due_x0020_Date: null,
      Client_x0020_Contact_x0020_A_x003: null,
      Client_x0020_Company: null,
      Status: null,
      Next_x0020_Contact_x0020_Date: null,
      NotesAdd: null,
      Comments: null,
      Title: null, //Project Owner
      isdataLoading: true,
      StatusOptions: [],
      ContactsDetails: [],
      CompanyDetails: [],
      Modified: null,
      hideDisplayDialogToRefresh: true,
    };
  }
  componentDidMount() {
    this.fetchOptionsFromchoiceColumn();
  }
  onTextFieldChange = (e, value) => {
    this.setState({ NotesAdd: value });
  };
  onCommentTextChange = (e, value) => {
    this.setState({ Comments: value });
  };
  AddMonth = () => {
    //let date = this.state.Next_x0020_Contact_x0020_Date;
    //date = moment(date).add(1, "months").toDate();
    this.setState({ Next_x0020_Contact_x0020_Date: null });
  };
  AddDays = () => {
    let date = this.state.Next_x0020_Contact_x0020_Date;
    date = moment(date).add(7, "days").toDate();
    this.setState({ Next_x0020_Contact_x0020_Date: date });
  };
  handleRefreshClick = () => {
    this.fetchSelecteItemDetails();
  };
  onSelectDate = (date) => {
    this.setState({ Next_x0020_Contact_x0020_Date: date });
  };
  onPrependNewNote = () => {
    const dateString = moment().format("MMM DD, YYYY");
    var _comments = dateString + " - " + this.state.NotesAdd + "\n" + this.state.Comments;
    this.setState({ Comments: _comments });
  };
  fetchOptionsFromchoiceColumn = () => {
    try {
      this.helper
        .getData(`/sites/${StaticConst.siteId}/lists/${StaticConst.lists.TCGProjectRegistry}/Columns/Status`)
        .then((res) => {
          let _choiceOptions = [];
          if (res.data.choice && res.data.choice.choices.length > 0) {
            res.data.choice.choices.forEach((element) => {
              _choiceOptions.push({ key: element.replace(/ /g, ""), text: element });
            });
          }
          this.setState(
            {
              StatusOptions: _choiceOptions,
            },
            () => {
              this.fetchSelecteItemDetails();
            }
          );
        })
        .catch((e) => {
          //this.setState({});
        });
    } catch (ex) {
      throw new Error(ex);
    }
  };
  getLookUpValue = (item) => {
    return item ? item.LookupValue : null;
  };
  grabAllContactsDetails = (item) => {
    var contacts = [];
    if (item["Client_x0020_Contact_x0020_A_x007"].length > 0) {
      item["Client_x0020_Contact_x0020_A_x007"].forEach((element, i) => {
        contacts.push({
          FullName: this.getLookUpValue(item["Client_x0020_Contact_x0020_A"][i]),
          Address: this.getLookUpValue(item["Client_x0020_Contact_x0020_A_x006"][i]),
          BusinessPhone: this.getLookUpValue(item["Client_x0020_Contact_x0020_A_x003"][i]),
          City: this.getLookUpValue(item["Client_x0020_Contact_x0020_A_x002"][i]),
          Company: this.getLookUpValue(item["Client_x0020_Contact_x0020_A_x00"][i]),
          CompanyContact: this.getLookUpValue(item["Client_x0020_Contact_x0020_A_x000"][i]),
          EmailAddress: this.getLookUpValue(item["Client_x0020_Contact_x0020_A_x001"][i]),
          ID: this.getLookUpValue(item["Client_x0020_Contact_x0020_A_x007"][i]),
          MobileNumber: this.getLookUpValue(item["Client_x0020_Contact_x0020_A_x004"][i]),
          ClientContactA_State_Province: this.getLookUpValue(item["Client_x0020_Contact_x0020_A_x005"][i]),
        });
      });
    }
    return contacts;
  };
  fetchSelecteItemDetails = () => {
    this.setState(
      {
        fields: null,
        isdataLoading: true,
      },
      () => {
        try {
          this.helper
            .getData(
              `/sites/${StaticConst.siteId}/lists/${StaticConst.lists.TCGProjectRegistry}/items/${this.props.location.selectedItemFields.id}?$expand=fields`
            )
            .then((res) => {
              const _fldItems = res.data.fields;
              this.setState({
                fields: res.data.fields,
                isdataLoading: false,
                Project_x0020_Name: _fldItems["Project_x0020_Name"],
                Bid_x0020__x0023_: _fldItems["Bid_x0020__x0023_"], //Bid #
                Estimated_x0020_Project_x0020_Va: _fldItems["Estimated_x0020_Project_x0020_Va"],
                Bid_x0020_Due_x0020_Date: _fldItems["Bid_x0020_Due_x0020_Date"]
                  ? moment(_fldItems["Bid_x0020_Due_x0020_Date"]).format("MM/DD/YYYY")
                  : null,
                Client_x0020_Contact_x0020_A_x003: _fldItems["Client_x0020_Contact_x0020_A_x003"],
                Client_x0020_Company: _fldItems["Client_x0020_Company"],
                Status: _fldItems["Status"],
                Next_x0020_Contact_x0020_Date: _fldItems["Next_x0020_Contact_x0020_Date"]
                  ? moment(_fldItems["Next_x0020_Contact_x0020_Date"]).toDate()
                  : null,
                NotesAdd: Office.context.mailbox.userProfile.displayName + ": ", //_fldItems["NotesAdd"],
                Comments: _fldItems["Comments"],
                Title: _fldItems["Title"], //Project Owner
                ContactsDetails: this.grabAllContactsDetails(_fldItems),
                Modified: _fldItems["Modified"],
                hideDisplayDialogToRefresh: true,
              });
            })
            .catch((e) => {
              this.setState({
                text: "res.data.fields.Title",
                successMessage: "",
              });
            });
        } catch (ex) {
          throw new Error(ex);
        }
      }
    );
  };
  onStatusChange = (event, option, index) => {
    this.setState({ Status: option.text });
  };
  onSaveClick() {
    this.helper
      .getData(
        `/sites/${StaticConst.siteId}/lists/${StaticConst.lists.TCGProjectRegistry}/items/${this.props.location.selectedItemFields.id}?$expand=fields`
      )
      .then((res) => {
        const _fldItems = res.data.fields;
        let _latestModified = moment(_fldItems["Modified"]);
        let _currentModified = moment(this.state.Modified);
        if (_latestModified.isSame(_currentModified)) {
          let param = {
            Comments: this.state["Comments"],
            Next_x0020_Contact_x0020_Date: this.state["Next_x0020_Contact_x0020_Date"]
              ? moment(this.state["Next_x0020_Contact_x0020_Date"]).format("LLL")
              : null,
            Status: this.state.Status,
          };
          this.setState(
            {
              isdataLoading: true,
            },
            () => {
              try {
                this.helper
                  .postData(
                    `/sites/${StaticConst.siteId}/lists/${StaticConst.lists.TCGProjectRegistry}/items/${this.props.location.selectedItemFields.id}/fields`,
                    param
                  )
                  .then((res) => {
                    this.handleRefreshClick();
                  })
                  .catch((e) => {});
              } catch (ex) {
                throw new Error(ex);
              }
            }
          );
        } else {
          this.setState({
            hideDisplayDialogToRefresh: false,
          });
        }
      });
  }
  render() {
    const styles = { root: { display: "inline-block" } };
    const calloutProps = { gapSpace: 0 };
    if (this.state.isdataLoading) {
      return <center><Label>Please wait</Label><Spinner message="" label="Getting Data from SharePoint List..." /></center>;
    }
    function DisplayItemsFromMultiLookUp({ objArray, displayName, forCompany }) {
      if (forCompany) {
        if (objArray && objArray.length > 0) {
          const html = objArray.map((item, index) => (
            <div key={index} className="ms-Grid-col ms-sm12">
              <TooltipHost
                tooltipProps={{
                  onRenderContent: () => (
                    <ul>
                      <li>Company Name: {item.LookupValue}</li>
                    </ul>
                  ),
                }}
                delay={TooltipDelay.zero}
                id={item.LookupId}
                directionalHint={DirectionalHint.topCenter}
                calloutProps={calloutProps}
                styles={styles}
              >
                <Link href="#">
                  <small aria-describedby={item.LookupId}> {item.LookupValue}</small>
                </Link>
              </TooltipHost>
            </div>
          ));
          return (
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12">
                <b>{displayName}</b>
              </div>
              {html}
            </div>
          );
        } else {
          return "";
        }
      } else {
        if (objArray && objArray.length > 0) {
          const html = objArray.map((item, index) => (
            <div key={index} className="ms-Grid-col ms-sm12">
              <TooltipHost
                tooltipProps={{
                  onRenderContent: () => (
                    <ul>
                      <li>Name: {item.FullName}</li>
                      {item.Company && <li>Company: {item.Company}</li>}
                      {item.CompanyContact && <li>Company Contact: {item.CompanyContact}</li>}
                      <li>Email Address: {item.EmailAddress}</li>
                      {item.BusinessPhone && <li>Business Phone: {item.BusinessPhone}</li>}
                      {item.MobileNumber && <li>Mobile Number: {item.MobileNumber}</li>}
                      {item.Address && <li>Address: {item.Address}</li>}
                    </ul>
                  ),
                }}
                delay={TooltipDelay.zero}
                id={item.ID}
                directionalHint={DirectionalHint.topCenter}
                calloutProps={calloutProps}
                styles={styles}
              >
                <Link href={`mailto:${item.EmailAddress}`}>
                  <small aria-describedby={item.ID}> {item.FullName}</small>
                </Link>
              </TooltipHost>
            </div>
          ));
          return (
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12">
                <b>{displayName}</b>
              </div>
              {html}
            </div>
          );
        } else {
          return "";
        }
      }
    }
    return (
      <div>
        <div className="ms-Grid" dir="ltr">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">
              <MyRegistryDetailList callParentRefreshFn={this.handleRefreshClick} />
              <br />
              <br />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">
              <RouterLink to="/">My Registry</RouterLink>
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">{this.state.Bid_x0020__x0023_}</div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">{this.state.Project_x0020_Name}</div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm6">
              <b>Value</b>
            </div>
            <div className="ms-Grid-col ms-sm6">
              <b>Due Date</b>
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm6">$ { (Math.round(this.state.Estimated_x0020_Project_x0020_Va * 100) / 100).toLocaleString()} </div>
            <div className="ms-Grid-col ms-sm6">{this.state.Bid_x0020_Due_x0020_Date}</div>
          </div>
          <DisplayItemsFromMultiLookUp objArray={this.state.ContactsDetails} displayName="Client Contacts" />
          <DisplayItemsFromMultiLookUp
            objArray={this.state.Client_x0020_Company}
            displayName="Client Companies"
            forCompany={true}
          />

          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm6">
              <Link
                href={`https://thecrossinggroup.sharepoint.com/ttc_usa/estimating/BidFiles/Forms/AllItems.aspx?id=/ttc_usa/estimating/BidFiles/${this.state.Bid_x0020__x0023_} ${this.state.Project_x0020_Name}`}
                target="_blank"
              >
                Bid Files
              </Link>
            </div>
            <div className="ms-Grid-col ms-sm6">
              <Link
                href={`https://thecrossinggroup.sharepoint.com/ttc_usa/projmgmt/ProjectFiles/Forms/AllItems.aspx?id=/ttc_usa/projmgmt/ProjectFiles/${this.state.Bid_x0020__x0023_} ${this.state.Project_x0020_Name}`}
                target="_blank"
              >
                Project Files
              </Link>
            </div>
          </div>

          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">
              <Dropdown
                defaultSelectedKey={this.state.Status.replace(/ /g, "")}
                placeholder="Select an option"
                label="Project Status"
                onChange={this.onStatusChange}
                options={this.state.StatusOptions}
              />
            </div>
          </div>

          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">
              <b>Next Contact Date</b>
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm8">
              <DatePicker
                showMonthPickerAsOverlay={true}
                value={this.state.Next_x0020_Contact_x0020_Date}
                onSelectDate={this.onSelectDate}
                placeholder="Select a date..."
                ariaLabel="Select a date"
              />
            </div>
            <div className="ms-Grid-col ms-sm2">
              <ActionButton onClick={this.AddDays}>+7d</ActionButton>
            </div>
            <div className="ms-Grid-col ms-sm2">
              <ActionButton onClick={this.AddMonth}>clear</ActionButton>
            </div>
          </div>

          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">
              <TextField
                value={this.state.NotesAdd}
                onChange={this.onTextFieldChange}
                label="New Note"
                multiline
                rows={3}
              />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm10"></div>
            <div className="ms-Grid-col ms-sm10">
              <PrimaryButton text="Prepend Data and New Note" onClick={this.onPrependNewNote}></PrimaryButton>
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">
              {/* <p> {this.state.Comments}</p> */}
              <TextField
                label="Comments"
                value={this.state.Comments}
                multiline
                rows={4}
                onChange={this.onCommentTextChange}
              />
            </div>
          </div>

          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm5"></div>
            <div className="ms-Grid-col ms-sm3">
              <PrimaryButton text="Save" onClick={this.onSaveClick} />
            </div>
            <div className="ms-Grid-col ms-sm3">
              <RouterLink to="/">
                <DefaultButton text="Cancel" />
              </RouterLink>
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">
              Last Updated : {moment(this.state.Modified).format("LLL").toString()}
            </div>
          </div>
          <Dialog
            hidden={this.state.hideDisplayDialogToRefresh}
            //onDismiss={toggleHideDialog}
            dialogContentProps={{
              type: DialogType.normal,
              title: "Update Conflict",
              closeButtonAriaLabel: "Close",
              subText: "Someone has updated the item, please reload item and continue...",
            }}
            //modalProps={modalProps}
          >
            <DialogFooter>
              <PrimaryButton onClick={this.handleRefreshClick} text="Reload" />
              {/* <DefaultButton  text="Don't send" /> */}
            </DialogFooter>
          </Dialog>
        </div>
      </div>
    );
  }
}
