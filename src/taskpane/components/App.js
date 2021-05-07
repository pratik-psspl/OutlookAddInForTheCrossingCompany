import * as React from "react";


import  ErrorHandling  from '../components/ErrorHandling';
import { ViewListItemDetails } from "./ViewListItemDetails";
import {CurrentUserDetails} from './CurrentUserDetails';
import { Spinner } from "office-ui-fabric-react";
import { HashRouter, Switch, Route } from "react-router-dom";
import { ViewAllListItems } from "./ViewAllListItems";
import { StaticConst } from "../helper/Const";
//import { PrimaryButton, DefaultButton } from "@fluentui/react";
/* global Button, Header, HeroList, HeroListItem, Progress */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {}
  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return <Spinner
      label={`Please Wait, Authenticating user with Site...`}
    />
    }

    return (
      <div className="ms-welcome">
        <ErrorHandling>
          <HashRouter>
            <Switch>
              <Route
                exact={true}
                path="/"
                render={() => <ViewAllListItems Authorization={this.props.Authorization} />}
              />
              <Route
                path="/ViewListItemDetails"
                render={(props) => <ViewListItemDetails {...props} Authorization={this.props.Authorization} />}
              />
              {/* <Route path='/OneDriveGridList' render={(props) => <OneDriveGridList {...props} sdkHelper={this.props.authenticator} />} /> */}
            </Switch>
          </HashRouter>
        </ErrorHandling>

        <CurrentUserDetails Authorization={this.props.Authorization}></CurrentUserDetails>
        <button className="clearCache" style={{display:"none"}}
         onClick={()=>{
          Office.context.roamingSettings.remove(StaticConst.roamingStorageName);
          Office.context.roamingSettings.saveAsync();
        }}>Log Out</button>
      </div>
    );
  }
}
