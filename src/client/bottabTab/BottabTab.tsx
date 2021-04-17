import * as React from "react";
import {
  Provider,
  Flex,
  Text,
  Button,
  Header,
  List,
  Alert,
  WindowMaximizeIcon,
  ExclamationTriangleIcon,
  Label,
  Input,
  ToDoListIcon
} from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * Implementation of the bottab Tab content page
 */
export const BottabTab = () => {
  const [{ inTeams, theme, context }] = useTeams();
  const [entityId, setEntityId] = useState<string | undefined>();

  const [todoItems, setTodoItems] = useState<string[]>([]);
  const [newTodoValue, setNewTodoValue] = useState<string>("");
  useEffect(() => {
    if (inTeams === true) {
      microsoftTeams.appInitialization.notifySuccess();
    } else {
      setEntityId("Not in Microsoft Teams");
    }
  }, [inTeams]);

  useEffect(() => {
    if (context) {
      setEntityId(context.entityId);
    }
  }, [context]);

  const handleOnChanged = (event): void => {
    setNewTodoValue(event.target.value);
  };

  const handleOnClick = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const newTodoItems = todoItems;
    newTodoItems.push(newTodoValue);
    setTodoItems(newTodoItems);
    setNewTodoValue("");
  };
  /**
   * The render() method to create the UI of the tab
   */
  return (
    <Provider theme={theme}>
      <Flex
        fill={true}
        column
        styles={{
          padding: ".8rem 0 .8rem .5rem"
        }}
      >
        <Flex.Item>
          <Header content="This is your tab" />
        </Flex.Item>
        <Flex.Item>
          <div>
            <div>
              <Text content={entityId} />
            </div>

            <div>
              <Button onClick={() => alert("It worked!")}>
                A sample button
              </Button>
            </div>
          </div>
        </Flex.Item>
        <Flex.Item
          styles={{
            padding: ".8rem 0 .8rem .5rem"
          }}
        >
          <Text size="smaller" content="(C) Copyright Perper" />
        </Flex.Item>
      </Flex>
      <Flex column gap="gap.smaller">
        <Header content="This is your tab" />
        <Alert
          icon={<ExclamationTriangleIcon />}
          content={entityId}
          dismissible
        ></Alert>
        <Text content="These are your to-do items:" size="medium"></Text>
        <List selectable>
          {todoItems.map((todoItem, i) => (
            <List.Item
              media={<WindowMaximizeIcon outline />}
              content={todoItem}
              index={i}
            ></List.Item>
          ))}
        </List>
        <Flex gap="gap.medium">
          <Flex.Item grow>
            <Flex>
              <Label
                icon={<ToDoListIcon />}
                styles={{
                  background: "darkgray",
                  height: "auto",
                  padding: "0 15px"
                }}
              ></Label>
              <Flex.Item grow>
                <Input
                  placeholder="New todo item"
                  fluid
                  value={newTodoValue}
                  onChange={handleOnChanged}
                ></Input>
              </Flex.Item>
            </Flex>
          </Flex.Item>
          <Button content="Add Todo" primary onClick={handleOnClick}></Button>
        </Flex>
        <Text content="(C) Copyright Contoso" size="smallest"></Text>
      </Flex>
    </Provider>
  );
};
