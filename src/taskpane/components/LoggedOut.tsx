import * as React from "react";
import { DefaultButton } from "@fluentui/react";

interface LoggedOutProps {
  onLogin: () => void;
}

const LoggedOut: React.FC<LoggedOutProps> = ({ onLogin }) => (
  <div>
      Please log in to Wire:<br/>
      <br/>
      <DefaultButton onClick={onLogin}>Log in</DefaultButton>
  </div>
);

export default LoggedOut;
