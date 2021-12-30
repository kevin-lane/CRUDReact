import * as React from 'react';
import "@pnp/sp/webs";
import styles from './CrudReact.module.scss';
import { ICrudReactProps } from './ICrudReactProps';
import { List } from './List/List';

export default class CrudReact extends React.Component<ICrudReactProps, {}> {
  public render(): React.ReactElement<ICrudReactProps> {
    return (
      <div className={ styles.crudReact }>
        <h1>CRUD React</h1>
        <List context={this.props.context} />
      </div>
    );
  }
}
