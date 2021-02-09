import {
	BaseClientSideWebPart,
	IPropertyPaneConfiguration,
	PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'InitiativeFormWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IInitiativeFormProps } from './components/IInitiativeFormProps';
import { InitiativeForm } from './components/InitiativeForm';


export interface IInitiativeFormWebPartProps {
	description: string;
}

export default class InitiativeFormWebPart extends BaseClientSideWebPart<IInitiativeFormWebPartProps> {

	public render(): void {
		const element: React.ReactElement<IInitiativeFormProps> = React.createElement(
			InitiativeForm,
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
