import * as React from "react";
import { Button } from "@fluentui/react-components";

interface LoggedOutProps {
  onLogin: () => void;
}

const LoggedOut: React.FC<LoggedOutProps> = ({ onLogin }) => (
  <div>
      Please log in to Wire:<br/>
      <br/>
      <Button onClick={onLogin}>Log in</Button>
  </div>
);

export default LoggedOut;
