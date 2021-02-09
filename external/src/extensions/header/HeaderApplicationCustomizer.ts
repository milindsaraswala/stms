import { override } from "@microsoft/decorators"
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from "@microsoft/sp-application-base"
import * as React from "react"
import * as ReactDOM from "react-dom"
import Header, { IHeaderProps } from "../../components/Header"

const LOG_SOURCE = "HeaderApplicationCustomizer"

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHeaderApplicationCustomizerProperties {
	// This is an example; replace with your own property
	testMessage: string
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HeaderApplicationCustomizer extends BaseApplicationCustomizer<
	IHeaderApplicationCustomizerProperties
> {
	private _topPlaceholder: PlaceholderContent | undefined

	@override
	public onInit(): Promise<void> {
		this._renderPlaceholders()

		return Promise.resolve()
	}

	private _onDispose(): void {
		console.log("Top Header dispose")
	}

	private _renderPlaceholders(): void {
		if (!this._topPlaceholder) {
			this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
				onDispose: this._onDispose,
			})

			if (!this._topPlaceholder) {
				console.error("The expected placeholder Top not found")
				return
			}

			const elem: React.ReactElement<IHeaderProps> = React.createElement(Header, {
				context: this.context,
			})
			if (this._topPlaceholder.domElement) {
				ReactDOM.render(elem, this._topPlaceholder.domElement)
			}
		}
	}
}
