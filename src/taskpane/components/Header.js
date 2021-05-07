import * as React from "react";

import * as OfficeHelpers from "@microsoft/office-js-helpers";
import { Authenticator, Storage, DefaultEndpoints } from "@microsoft/office-js-helpers";
export default class Header extends React.Component {
  constructor(props) {
    super(props);
    this.state = {};
  }
  componentDidMount() {
    var authenticator = new OfficeHelpers.Authenticator();
    if (OfficeHelpers.Authenticator.isAuthDialog()) return;
    authenticator.endpoints.registerMicrosoftAuth("1fbf4dfc-081e-4823-9aec-8f88f6616382", {
      responseType: "code",
      redirectUrl: "https://localhost:3000/taskpane.html",
      scope: "openid profile offline_access User.Read",
    });
    
    authenticator = new OfficeHelpers.Authenticator();
    authenticator
      .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft, false)
      .then(function (token) {
        callback(true);
        _this.login(token["code"]);
      })
      .catch(function (e) {
        callback(false);
        OfficeHelpers.Utilities.log(e, "Error on authentication");
      });
  }
  render() {
    const { title, logo, message } = this.props;

    return (
      <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
        <h1>prat</h1>
      </section>
    );
  }
}
