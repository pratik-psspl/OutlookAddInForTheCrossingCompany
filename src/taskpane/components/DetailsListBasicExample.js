import * as React from "react";
import { Announced } from "office-ui-fabric-react/lib/Announced";
import { TextField, ITextFieldStyles } from "office-ui-fabric-react/lib/TextField";
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import { Text } from "office-ui-fabric-react/lib/Text";
import axios from "axios";

const exampleChildClass = mergeStyles({
  display: "block",
  marginBottom: "10px",
});

export class DetailsListBasicExample extends React.Component {
  constructor(props) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    // Populate with items for demos.
    this._columns = [
      { key: "column1", name: "name", fieldName: "name", minWidth: 100, maxWidth: 200, isResizable: true },
      { key: "column2", name: "webUrl", fieldName: "webUrl", minWidth: 100, maxWidth: 200, isResizable: true },
    ];

    this.state = {
      items: [],
      allItems:[],
      selectionDetails: this._getSelectionDetails(),
    };
  }
  componentDidMount() {
    axios
      .get(`https://graph.microsoft.com/v1.0/sites?search=*`, {
        headers:{ Authorization: this.props.Authorization } ,
      })
      .then((res) => {
        console.log(res);
        var a = res.data.value;
        this.setState({
            items:a,
            allItems:a
        })    
      })   
  }
  render() {
    const { items, selectionDetails } = this.state;

    return (
      <Fabric>
        <TextField className={exampleChildClass} label="Filter by name:" onChange={this._onFilter} />
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={items}
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
          />
        </MarqueeSelection>
      </Fabric>
    );
  }

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
}
