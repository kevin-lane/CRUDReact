import * as React from 'react';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react';
import { Panel } from '@fluentui/react/lib/Panel';
import { useBoolean } from '@fluentui/react-hooks';

export function SidePanel(props) {
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(false);
  const [titleInput, setTitleInput] = React.useState('');
  const [costInput, setCostInput] = React.useState('');
  
  async function addItem(){
    if (titleInput.trim() == "") {
      alert('Please enter a title!');
    }
    else {
      const iar: IItemAddResult = await sp.web.lists.getByTitle("Tulips").items.add({
        Title: titleInput.trim(),
        ManufacturingCost: costInput,
      });
      console.log(iar.data);
      
      alert(`New item has been added`);
      dismissPanel();
    }
  }

  async function updateItem(id) {
    id = props.itemToUpdateID;
    var prevTitle = props.itemTitle;
    var prevCost = props.itemCost;

    let list = sp.web.lists.getByTitle("Tulips");
    await list.items.getById(id).update({
      Title: titleInput == "" ? prevTitle : titleInput.trim(),
      ManufacturingCost: costInput == "" ? prevCost : costInput
    });
    dismissPanel();
    alert(`Item with id ${id} has been updated!`);
  }
  
  return (
    <div>
      <DefaultButton text={props.sidePanelButtonText} onClick={openPanel} disabled={props.buttonDisabled} />
      <Panel
        isOpen={isOpen}
        hasCloseButton={false} 
        headerText={props.toUpdateItem ? `Updating item with ID ${props.itemToUpdateID}` : "Add a new list item"}
      >
      <TextField 
        label="Title" 
        defaultValue={props.toUpdateItem ? props.itemTitle : ''}
        onChange={(e) => setTitleInput((e.target as HTMLInputElement).value)} 
        required 
      />
      <br/>
      <TextField 
        label='Manufacturing Cost'
        type='number'
        min={0}
        defaultValue={props.toUpdateItem ? props.itemCost : '0'}
        onChange={(e) => setCostInput((e.target as HTMLInputElement).value)}
      />
      <br/>
        <PrimaryButton onClick={props.toUpdateItem ? updateItem : addItem } text={props.actionButtonText}/>
        <DefaultButton onClick={dismissPanel} text="Close panel" />
      </Panel>
    </div>
  );
}