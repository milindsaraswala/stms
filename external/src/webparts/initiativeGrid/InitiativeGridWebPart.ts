import {
	BaseClientSideWebPart,
	IPropertyPaneConfiguration,
	PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'InitiativeGridWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IInitiativeGridProps } from './components/IInitiativeGridProps';
import { InitiativeGrid } from './components/InitiativeGrid';

export interface IInitiativeGridWebPartProps {
	description: string;
}

export default class InitiativeGridWebPart extends BaseClientSideWebPart<IInitiativeGridWebPartProps> {

	public render(): void {
		const element: React.ReactElement<IInitiativeGridProps> = React.createElement(
			InitiativeGrid,
			{
				description: this.properties.description,
				context: this.context
			}
		);

		ReactDom.render(element, this.domElement);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneTextField('description', {
									label: strings.DescriptionFieldLabel
								})
							]
						}
					]
				}
			]
		};
	}
}
