import * as React from "react";
import { Button } from "@fluentui/react-components";
import type { SelfUser } from "../../types/SelfUser";

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
      <Button onClick={onLogout}>Disconnect Add-in</Button>
  </div>
);

export default LoggedIn;
