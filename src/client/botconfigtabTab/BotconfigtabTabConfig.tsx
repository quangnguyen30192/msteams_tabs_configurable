import * as React from "react";
import {
  Provider,
  Input,
  DropdownProps,
  Dropdown,
  Flex,
  Text,
  Button,
  Header
} from "@fluentui/react-northstar";
import { useState, useEffect, useRef } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * Implementation of botconfigtab configuration page
 */
export const BotconfigtabTabConfig = () => {
  const [{ inTeams, theme, context }] = useTeams({});
  const [text, setText] = useState<string>();
  const [mathOperator, setMathOperator] = useState<string>();
  const entityId = useRef("");

  const onSaveHandler = (saveEvent: microsoftTeams.settings.SaveEvent) => {
    const host = "https://" + window.location.host;
    microsoftTeams.settings.setSettings({
      contentUrl:
        host +
        "/botconfigtabTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}",
      websiteUrl:
        host +
        "/botconfigtabTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}",
      suggestedDisplayName: "botconfigtab",
      removeUrl: host + "/botconfigtabTab/remove.html?theme={theme}",
      entityId: entityId.current
    });
    saveEvent.notifySuccess();
  };

  useEffect(() => {
    if (context) {
      setText(context.entityId);
      setMathOperator(context.entityId.replace("MathPage", ""));
      entityId.current = context.entityId;
      microsoftTeams.settings.registerOnSaveHandler(onSaveHandler);
      microsoftTeams.settings.setValidityState(true);
      microsoftTeams.appInitialization.notifySuccess();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [context]);

  return (
    <Provider theme={theme}>
      <Flex fill={true}>
        <Flex.Item>
          <div>
            <Header content="Configure your tab" />
            <Input
              placeholder="Enter a value here"
              fluid
              clearable
              value={text}
              onChange={(e, data) => {
                if (data) {
                  setText(data.value);
                  entityId.current = data.value;
                }
              }}
              required
            />
          </div>
        </Flex.Item>
      </Flex>{" "}
      <Flex gap="gap.smaller" style={{ height: "300px" }}>
        <Dropdown
          placeholder="Select the math operator"
          items={["add", "subtract", "multiply", "divide"]}
          onChange={(e, data) => {
            if (data) {
              let op = data.value ? data.value.toString() : "add";
              setMathOperator(op);
              entityId.current = `${op}MathPage`;
            }
          }}
          value={mathOperator}
        ></Dropdown>
      </Flex>
    </Provider>
  );
};
