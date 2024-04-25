import * as React from "react";
import Progress from "./Progress";
import LoggedIn from "./LoggedIn";
import LoggedOut from "./LoggedOut";
import { authorizeDialog, revokeOauthToken } from "../../wireAuthorize/wireAuthorize";
import { removeTokens } from "../../utils/tokenStore";
import { setUserDetails, removeUserDetails, getUserDetails } from "../../utils/userDetailsStore";
import { getSelf } from "../../calendarIntegration/getSelf";
import { SelfUser } from "../../types/SelfUser";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  isLoggedIn: boolean;
  user: SelfUser | null;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      isLoggedIn: false,
      user: null,
    };
  }

  componentDidMount() {
    const user = getUserDetails();
    if (user) {
      this.setState({ isLoggedIn: true, user });
    }
  }

  login = async () => {
    const isLoggedIn = await authorizeDialog();
    this.setState({ isLoggedIn });
    if (isLoggedIn) {
      const user = await getUserDetails();
      this.setState({ isLoggedIn, user });
    }
  };

  logout = async () => {
    await revokeOauthToken();
    removeTokens();
    removeUserDetails();
    this.setState({ isLoggedIn: false, user: null });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;
    const { isLoggedIn, user } = this.state;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo={require("./../../../assets/icon-128.png")} message="Loading in progress..." />
      );
    }

    return (
      <div className="ms-Grid">
      <div className="ms-Grid-row">
        <div className="ms-Grid-col">
          <h1>Settings</h1>
  
        {isLoggedIn && user ? <LoggedIn user={user} onLogout={this.logout} /> : <LoggedOut onLogin={this.login} />}
      </div>
      </div>
      </div>
    );
  }
}
