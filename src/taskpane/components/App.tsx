import { UserAgentApplication } from 'msal';
import { getUserDetails } from '../../services/GraphService';
import 'bootstrap/dist/css/bootstrap.css';
// import Calendar from './Calendar';
import React = require('react');
import { config } from './Config';
import { Button } from 'reactstrap';
import ExcelImages from './ExcelImages';

export interface AppState {
  error: any,
  isAuthenticated: boolean;
  user: any
}

class App extends React.Component<{}, AppState> {
  userAgentApplication: any
  constructor(props) {
    super(props);
    this.login = this.login.bind(this);
    this.logout = this.logout.bind(this);
    this.userAgentApplication = new UserAgentApplication({
        auth: {
            clientId: config.appId,
            redirectUri: config.redirectUri
        },
        cache: {
            cacheLocation: "localStorage",
            storeAuthStateInCookie: true
        }
    });

    var user = this.userAgentApplication.getAccount();

    this.state = {
      isAuthenticated: (user !== null),
      user: {},
      error: null
    };

    if (user) {
      // Enhance user object with data from Graph
      this.getUserProfile();
    }
  }

  render() {
    return (
        <div>
          {!this.state.isAuthenticated ?
          <Button onClick={this.login}>Sign in</Button>
          : <Button onClick={this.logout}>Sign out</Button>}
          {this.state.isAuthenticated ? 
          <span>
            Welcome: {this.state.user.displayName}
            {/* <Calendar />  */}
            <ExcelImages />
          </span> 
          : <span></span>}
        </div>
    );
  }

  async login() {
    try {
      await this.userAgentApplication.loginPopup(
        {
          scopes: config.scopes,
          prompt: "select_account"
      });
      await this.getUserProfile();
    }
    catch(err) {
    }
    //let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true });
    // console.log(bootstrapToken);
  }

  logout() {
    this.userAgentApplication.logout();
  }

  async getUserProfile() {
    try {
      // Get the access token silently
      // If the cache contains a non-expired token, this function
      // will just return the cached token. Otherwise, it will
      // make a request to the Azure OAuth endpoint to get a token

      var accessToken = await this.userAgentApplication.acquireTokenSilent({
        scopes: config.scopes
      });

      if (accessToken) {
        // Get the user's profile from Graph
        var user = await getUserDetails(accessToken);
        this.setState({
          isAuthenticated: true,
          user: {
            displayName: user.displayName,
            email: user.mail || user.userPrincipalName
          },
          error: null
        });
      }
    }
    catch(err) {
    }
  }
}

export default App;
