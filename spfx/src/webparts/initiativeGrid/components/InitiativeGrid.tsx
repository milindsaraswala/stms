import * as React from 'react';
import styles from './InitiativeGrid.module.scss';
import { IInitiativeGridProps } from './IInitiativeGridProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class InitiativeGrid extends React.Component < IInitiativeGridProps, {} > {
  public render(): React.ReactElement<IInitiativeGridProps> {
    return(
      <div className = { styles.initiativeGrid } >
  <div className={styles.container}>
    <div className={styles.row}>
      <div className={styles.column}>
        <span className={styles.title}>Welcome to SharePoint!</span>
        <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
        <p className={styles.description}>{escape(this.props.description)}</p>
        <a href='https://aka.ms/spfx' className={styles.button}>
          <span className={styles.label}>Learn more</span>
        </a>
      </div>
    </div>
  </div>
      </div >
    );
  }
}
