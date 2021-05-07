import * as React from "react";
import { Fabric ,Link,IconButton} from "office-ui-fabric-react";
import { Link as RouterLink } from "react-router-dom";

export class MyRegistryDetailList extends React.Component {
  constructor(props) {
    super(props);
    this.state = {};
    this.callParentRefreshFn=this.callParentRefreshFn.bind(this);
  }
  callParentRefreshFn(){
    this.props.callParentRefreshFn();
  }
 
  render() {
    return (
        <div className="ms-Grid" dir="ltr">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm4">
              <Link href="https://thvc.sharepoint.com/BusDev/Lists/Leads%20List/AllItems.aspx" target="_blank">
                <small>Open Registry</small>
              </Link>
            </div>
            <div className="ms-Grid-col ms-sm3">
              <Link href={`https://thvc.sharepoint.com/busdev/sitepages/iw_NewForm.aspx?pageType=8&lID=85e8b9c6-6be0-4f87-948c-2d798c49de1f`} target="_blank">
                <small>New Bid</small>
              </Link>
            </div>
            <div className="ms-Grid-col ms-sm4">
              <Link href={`https://thvc.sharepoint.com/BusDev/Lists/Customer%20Contacts/AllItems.aspx?FilterField1=Email&FilterValue1=${Office.context.mailbox.userProfile.emailAddress}&FilterType1=Text`} target="_blank">
                <small>Find Contact</small>
              </Link>
            </div>
            <div className="ms-Grid-col ms-sm1">             
                <IconButton style={{margin:"-4px ​0px 0px -19p"}} iconProps={{ iconName: "Refresh" }} title="Refresh" ariaLabel="Refresh" onClick={this.callParentRefreshFn}/>            
            </div>
          </div>
        </div>      
    );
  }
}