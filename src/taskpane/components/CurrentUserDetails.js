import * as React from "react";
import { AsyncHelper } from "../helper/AsyncHelper";
import { Persona, PersonaSize, Sticky, StickyPositionType } from "office-ui-fabric-react";
export class CurrentUserDetails extends React.Component {
  helper = new AsyncHelper(this.props.Authorization);
  constructor(props) {
    super(props);
    this.state = {
      displayName: "",
      mail: "",
      displayUser: false,
    };
  }
  componentDidMount() {
    console.log(Office.context.mailbox.userProfile.emailAddress);

    this.helper.getData(`/me`).then((res) => {
      this.setState({
        displayName: res.data.displayName,
        mail: res.data.mail,
        displayUser: true,
      });
    });
  }
  render() {
    const examplePersona = {
      text: this.state.displayName,
      secondaryText: this.state.mail,
      showSecondaryText: true,
    };
    return (
      this.state.displayUser && (
        <div className="contactDetailsFooter">
          <Persona {...examplePersona} size={PersonaSize.size24} imageAlt="Annie Ried, status is unknown" />
        </div>
      )
    );
  }
}
