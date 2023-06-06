import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import { SelfUser } from "../../types/SelfUser";

interface LoggedInProps {
  user: SelfUser | null;
  onLogout: () => void;
}

const LoggedIn: React.FC<LoggedInProps> = ({ user, onLogout }) => (
  <div>
      Welcome!<br/>
      <br/>
      Display Name: {user.name}<br/> 
      Username: {user.handle}<br/>
      E-mail: {user.email}<br/>
      <br/>
      <DefaultButton onClick={onLogout}>Log out</DefaultButton>
  </div>
);

export default LoggedIn;
