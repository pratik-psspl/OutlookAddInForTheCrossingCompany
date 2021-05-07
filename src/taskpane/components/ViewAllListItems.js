import * as React from "react";
import { Announced } from "office-ui-fabric-react/lib/Announced";
import { TextField, ITextFieldStyles } from "office-ui-fabric-react/lib/TextField";
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import { Text } from "office-ui-fabric-react/lib/Text";
import { Spinner, Label, SwatchColorPicker } from "office-ui-fabric-react";
import { Link } from "react-router-dom";
import { IOverflowSetItemProps, OverflowSet } from "office-ui-fabric-react/lib/OverflowSet";
import { MyRegistryDetailList } from "./MyRegistryDetailList";
import axios from "axios";
import { StaticConst } from "../helper/Const";
import { AsyncHelper } from "../helper/AsyncHelper";
//import { _ } from "core-js";
import _ from "lodash";
import moment from "moment";
const colorCellsExample2 = [{ id: "a", label: "red", color: "#a4262c" }];
export class ViewAllListItems extends React.Component {
  helper = new AsyncHelper(this.props.Authorization);
  constructor(props) {
    super(props);
    this.state = {
      items: [],
      allItems: [],
      error: null,
      isLoading: true,
    };

    // Populate with items for demos.
    this._columns = [{ key: "id", name: "id", fieldName: "id", minWidth: 100, maxWidth: 200, isResizable: true }];
    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });
  }

  handleRefreshClick = () => {
    this.fetchAllItemsFromList();
  };
  componentDidMount() {
    this.fetchAllItemsFromList();
  }
  fetchAllItemsFromList = () => {
    this.setState(
      {
        items: [],
        allItems: [],
        isLoading: true,
        error: null,
      },
      () => {
        try {
          this.helper
            .getData(`/sites/${StaticConst.siteId}/lists/${StaticConst.lists.TCGProjectRegistry}/items?$expand=fields`)
            .then((res) => {              
              this.setState({
                items: this.filterListItems(
                  res.data.value,
                  StaticConst.currentUser ? StaticConst.currentUser : Office.context.mailbox.userProfile.emailAddress
                ),
                allItems: res.data.value,
                isLoading: false,
                error: null,
              });
            })
            .catch((e) => {
              this.setState({
                error: e,
                isLoading: false,
              });
            });
        } catch (ex) {
          throw new Error(ex);
        }
      }
    );
  };
  filterListItems = (allItems, currentUserEmailId) => {
    let items = [];
    if (currentUserEmailId) {
      allItems.forEach((array, index) => {
        let item = array.fields;
        if (item && item["TCG_x0020_Opportunity_x0020_Mana"] && item["TCG_x0020_Opportunity_x0020_Mana"].length > 0) {
          let isCurrentUser = false;
          item["TCG_x0020_Opportunity_x0020_Mana"].forEach((element) => {
            if (element.Email === currentUserEmailId) {
              isCurrentUser = true;
            }
          });
          if (isCurrentUser === true) {
            items.push(array);
          }
        }
      });
    }
    
    var filteredByRecord= _.orderBy(
      items,
      [
        (item) => {
          if (item.fields.Next_x0020_Contact_x0020_Date) {
            return new Date(item.fields.Next_x0020_Contact_x0020_Date);
          }
          return new Date(0);
        },
      ],
      ["desc"]
    );

    return filteredByRecord;
  };
  _getSelectionDetails() {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return "No items selected";
      case 1:
        return "1 item selected: " + this._selection.getSelection()[0].name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  _onFilter = (ev, text) => {
    this.setState({
      items: text ? this.state.allItems.filter((i) => i.name.toLowerCase().indexOf(text) > -1) : this.state.allItems,
    });
  };

  _onItemInvoked = (item) => {
    alert(`Item invoked: ${item.name}`);
  };

  _renderItemColumn(item, index, column) {
    function ColorIconBasedOnStatus({ nextDate }) {
      
      var todaysDate = moment();
      var next5Days = moment().add(5, "days");
      var dateValue = moment(nextDate);
      let color = [];

      if (dateValue.isAfter(todaysDate) && dateValue.isBefore(next5Days)) {
        color = "yellow icon";
      } else if (dateValue.isSameOrBefore(todaysDate)) {
        //today or older
        color =  "red icon" ;
      } else {
        color = "blue icon";
      }

      return (
        <div
          
          className={color}
          
        />
      );
    }

    return (
      <React.Fragment>
        <div className="listing">
            <div className="icon-block">
              <ColorIconBasedOnStatus nextDate={item.fields.Next_x0020_Contact_x0020_Date}></ColorIconBasedOnStatus>
            </div>
            <div className="content-list">
                <div className="list-title">
                  <Link
                to={{ pathname: "/ViewListItemDetails", selectedItemFields: item.fields }}
                className="btn-primary full-width"
              >
                {item.fields.Bid_x0020__x0023_} {item.fields.Project_x0020_Name}
              </Link>
              </div>
              <div className="inner-content">
                  <p><label>Next :</label> <span>{moment(item.fields.Next_x0020_Contact_x0020_Date).format("YYYY-MM-DD")}</span></p>
                  <p><label>Status :</label> <span>{item.fields.Status}</span></p>
                  <p><label>Value :</label> <span>$ { (Math.round(item.fields.Estimated_x0020_Project_x0020_Va * 100) / 100).toLocaleString()}</span></p>
              </div>
            </div>
        </div>
        </React.Fragment>
    );
  }
  onRenderDetailsHeader(detailsHeaderProps) {
    return <p></p>;
  }
  render() {
    if (this.state.error) {
      throw new Error(this.state.error);
    }
    return (
      <div>
        <MyRegistryDetailList callParentRefreshFn={this.handleRefreshClick} />

        {this.state.isLoading == true && (
          <center>
            <div className="loding-block">
            <Label>Please wait</Label>
            <Spinner
              label={`Getting Details selected item...`}
            />
            </div>
          </center>
        )}

        {this.state.isLoading == false && this.state.error == null && (
          <div>
                <h3 class="title">My Registry with Next Contact Date</h3>
            <div className="ms-Grid" dir="ltr">
            <div className="ms-Grid-row content-block">
              {this.state.items.length<1 && <Label>No Record Found</Label>}
              <MarqueeSelection selection={this._selection}>
                <DetailsList
                  items={this.state.items}
                  columns={this._columns}
                  selectionMode={SelectionMode.none}
                  setKey="set"
                  layoutMode={DetailsListLayoutMode.justified}
                  selection={this._selection}
                  selectionPreservedOnEmptyClick={false}
                  ariaLabelForSelectionColumn="Toggle selection"
                  ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                  checkButtonAriaLabel="select row"
                  onItemInvoked={this._onItemInvoked}
                  onRenderItemColumn={this._renderItemColumn}
                  onRenderDetailsHeader={this.onRenderDetailsHeader}
                />
              </MarqueeSelection>
            </div>
          </div>
        </div>
        )}
      </div>
    );
  }
}
