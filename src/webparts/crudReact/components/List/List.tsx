import * as React from 'react';
import { ICrudReactProps } from '../ICrudReactProps';
import { DetailsList, IColumn, Selection } from '@fluentui/react/lib/DetailsList';
import { SPHttpClient, SPHttpClientResponse, MSGraphClient } from '@microsoft/sp-http';
import { IListItemProps } from './IListItemProps';
import { Icon } from '@fluentui/react';
import { sp } from "@pnp/sp";
import { DialogBox } from '../Dialog/DialogBox';
import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { SidePanel } from '../Panel/SidePanel';
import { graph } from '@pnp/graph';
import "@pnp/graph/users";
import { IMailMessage } from '../Mail/IMailMessage';

export interface IListItemState{
    Title: string;
    ManufacturingCost: number;
    Retail_x0020_Price: number;
    ID: number;
    ControlButtons: any;
}

export interface IListItems{
    items: IListItemState[];
    hideDialog: boolean;
    selectionDetails: any;
    toUpdate: boolean;
    itemToUpdateID: number;
    itemToUpdateTitle: string;
    itemToUpdateCost: string;
    toDisable: boolean;
}

export class List extends React.Component<IListItemProps, IListItems> {
    private itemSelection: Selection;
    private listColumns: IColumn[];
    private listItems: IListItemState[];
    constructor(props: ICrudReactProps) {
        super(props);

        this.listItems = [];
        this.listColumns = [
            { key: 'column1', name: 'Title', fieldName: 'Title', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column2', name: 'Manufacturing Cost', fieldName: 'ManufacturingCost', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column3', name: 'Retail Price', fieldName: 'Retail_x0020_Price', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'column4', name: '', fieldName: 'ControlButtons', minWidth: 100, maxWidth: 200, isResizable: true }
        ];

        this.itemSelection = new Selection({
            onSelectionChanged: () => this.setState({ selectionDetails: this.getSelectionDetails() })
        });

        this.state = {
            items: this.listItems,
            hideDialog: true,
            selectionDetails: this.getSelectionDetails(),
            toUpdate: false,
            itemToUpdateID: null,
            itemToUpdateTitle: '',
            itemToUpdateCost: '',
            toDisable: false
        };
    }

    public componentDidMount(): void {
        console.log("componentDidMount Called");
        this._renderList();
    }

    public componentDidUpdate(prevProps: Readonly<IListItemProps>, prevState: Readonly<IListItems>, snapshot?: any): void {
        if (this.state.items.length !== prevState.items.length) {
            console.log("Updated");
            this._getListData();
        }
    }

    public _getListData() : Promise<IListItemState[]>{
        return this.props.context.spHttpClient.get(
            this.props.context.pageContext.web.absoluteUrl + 
            "/_api/web/lists/GetByTitle('Tulips')/Items", SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json().then(data => {
                    return data.value;
                });
            });
    }

    public _renderList(){
        this._getListData().then(response => {
            response.forEach(item => {
                this.state.items.push({
                    Title: item.Title,
                    ManufacturingCost: item.ManufacturingCost,
                    Retail_x0020_Price: item.Retail_x0020_Price,
                    ID: item.ID,
                    ControlButtons: <Icon iconName='Delete' onClick={() => this._deleteListItem(item.ID)} />
                });
            });
            this.setState({
                items: this.state.items
            });
        });        
    }

    private getSelectionDetails(){
        const selectionCount = this.itemSelection.getSelectedCount();
        const selectedItem = (this.itemSelection.getSelection() as any);
        selectedItem.map(res => {
            this.setState({
                itemToUpdateID: res.ID,
                itemToUpdateTitle: res.Title,
                itemToUpdateCost: res.ManufacturingCost
            });
        });

        if (selectionCount > 1) {
            //Disable update/add button if user chooses more than one selections of items
            this.setState({ toDisable: true });
        }
        else{
            this.setState({ toDisable: false });
            switch (selectionCount) {
                case 0:
                    this.setState({ toUpdate: false });
                    break;
                case 1:
                    this.setState({ toUpdate: true });
                    break;
            }
        }
    }

    public async _deleteListItem(id){
        let message = `Are you sure you want to delete list item ${id}?`;
        if (confirm(message) == true) {
            let list = sp.web.lists.getByTitle("Tulips");
            var currentUser = await graph.me();

            const deletedItem = await sp.web.lists.getByTitle("Tulips")
                .items
                .getById(id)
                .select("Id", "Title", "Author/ID", "Author/Title", "Author/EMail")
                .expand("Author")
                .get();

            await list.items.getById(id).delete().then();
            alert(`Item ${id} has been deleted by ${currentUser.displayName}!`);
            
            var itemsAfterDelete = this.state.items.filter(function(item) {
                return item.ID != deletedItem.ID;
            });
            
            this.setState({
                items: itemsAfterDelete
            });
            console.log(this.state.items);

            const deletionMessage: IMailMessage = {
                message: {
                    subject: `Item ${deletedItem.Title}(${deletedItem.ID}) has been deleted`,
                    body: {
                        contentType: "HTML",
                        content: `The list item ${deletedItem.Title} you created has been deleted by ${currentUser.displayName}!`
                    },
                    toRecipients: [
                        {
                            emailAddress: {
                                address: deletedItem.Author.EMail
                            }
                        },
                    ],
                },
                saveToSentItems: true
            };

            if (deletedItem.Author.EMail != currentUser.mail) {
                this.props.context.msGraphClientFactory.getClient()
                    .then((client: MSGraphClient) => {                        
                        client.api('/me/sendMail').post(deletionMessage);
                    });
            }
            else{
                return;
            }
            
        }
        else{
            return;
        }    
    }
    
    public render(): JSX.Element {
        const { items } = this.state;        

        return (
            <div>
                {this.state.toUpdate ? 
                    <SidePanel 
                        sidePanelButtonText="Update Item" 
                        actionButtonText="Update Item" 
                        toUpdateItem={this.state.toUpdate}
                        itemToUpdateID={this.state.itemToUpdateID}
                        itemTitle={this.state.itemToUpdateTitle}
                        itemCost={this.state.itemToUpdateCost}
                        buttonDisabled={this.state.toDisable} 
                    /> 
                    : 
                    <SidePanel 
                        sidePanelButtonText="Add Item" 
                        actionButtonText="Add Item" 
                        toUpdateItem={this.state.toUpdate}
                        buttonDisabled={this.state.toDisable} 
                    />
                }
                <DialogBox hideDialog={this.state.hideDialog} onConfirm={this._deleteListItem} onCancel={() => this.setState({ hideDialog: true })} />
                <MarqueeSelection selection={this.itemSelection}>
                    <DetailsList 
                        items={items}
                        columns={this.listColumns}
                        selection={this.itemSelection}
                    />
                </MarqueeSelection>
            </div>
        );
    }
}