import * as React from 'react';
// import styles from './Myexpense.module.scss';
import { IMyexpenseProps } from './IMyexpenseProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import Dropdowns from './expense';

export default class Myexpense extends React.Component<IMyexpenseProps, {}> {

  handleResourceSelectionChange = (selectedResources: any) => {
    // Implement logic to handle selected resources
    console.log('Selected Resources:', selectedResources);
  };

  handleTimePeriodChange = (selectedTimePeriod: any) => {
    // Implement logic to handle selected time period
    console.log('Selected Time Period:', selectedTimePeriod);
  };

  handleAllocationChange = (selectedAllocation: any) => {
    // Implement logic to handle selected allocation
    console.log('Selected Allocation:', selectedAllocation);
  };
  public render(): React.ReactElement {
    return (
      <div>
        <Dropdowns
          onResourceSelectionChange={this.handleResourceSelectionChange}
          onTimePeriodChange={this.handleTimePeriodChange}
          onAllocationChange={this.handleAllocationChange}
        />
      </div>
    );
  }
}
