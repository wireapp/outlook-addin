import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import { SelfUser } from "../../types/SelfUser";

interface LoggedInProps {
  user: SelfUser | null;
  onLogout: () => void;
}

const LoggedIn: React.FC<LoggedInProps> = ({ user, onLogout }) => (
  <div>
      <p>Welcome, {user.name}! You are logged in on 
      Wire for Outlook with the following credentials:
      </p>
      Username: {user.handle}<br/>
      E-mail: {user.email}<br/>
      <br/>
      <DefaultButton onClick={onLogout}>Disconnect Add-in</DefaultButton>
  </div>
);

export default LoggedIn;
