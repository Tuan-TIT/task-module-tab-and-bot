import React, { useEffect } from "react";
import { Welcome } from "./sample/Welcome";
import { Button } from "@fluentui/react-components";

export default function Tab() {

  const newWindowObject = window as any;

  useEffect(() => {
    newWindowObject.microsoftTeams.initialize();
  }, [])

  const submitTask = () => {
    newWindowObject.microsoftTeams.tasks.submitTask({ abc: 1 });
  }

  return (
    <div>
      <Welcome />
      <Button onClick={submitTask}>submit</Button>
    </div>
  );
}
