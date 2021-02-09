import { ExtensionContext } from "@microsoft/sp-extension-base"
import "@pnp/sp/items"
import "@pnp/sp/lists"
import "@pnp/sp/webs"
// import "devextreme/dist/css/dx.carmine.css"
// import "devextreme/dist/css/dx.common.css"
import React, { useState } from "react"
// import "../assets/css/boxicons.min.css"
// import "../assets/css/custom.css"

const logo: any = require("../assets/images/logo.png").default

export interface IHeaderProps {
	context: ExtensionContext
}

const dropDownButtonRender = (button) => {
	return (
		<div>
			<i className="bx bx-bell bx-tada"></i>
			<span className="badge badge-danger badge-pill">3</span>
		</div>
	)
}

const Header = (props: IHeaderProps) => {
	const [pendingItems, setPendingItems] = useState([])
	const [pendingItemsCount, setPendingItemsCount] = useState(0)

	// useEffect(() => {
	// 	sp.setup({
	// 		spfxContext: props.context,
	// 	})
	// 	getPendingItems()
	// }, [])

	// const getPendingItems = () => {
	// 	sp.web.lists
	// 		.getByTitle("Initiatives")
	// 		.items.select("Id", "Title", "Created")
	// 		.filter(`PendingWithId eq ${props.context.pageContext.legacyPageContext.userId}`)
	// 		.getAll()
	// 		.then((items) => {
	// 			setPendingItems(items)
	// 			setPendingItemsCount(items.length)
	// 		})
	// }

	const renderItems = (item) => {
		return (
			<div style={{ padding: "20px", width: "1500px" }}>
				<a>
					<h6>{item.Title}</h6>
				</a>
			</div>
		)
	}

	return (
		<React.Fragment>
			<header id="page-topbar">
				<div className="navbar-header">
					<div className="d-flex">
						<div className="navbar-brand-box">
							<a href="/" className="logo">
								<span>
									<img src={String(logo)} alt="" height="17" />
								</span>
							</a>
						</div>
					</div>
					<div className="d-flex">
						<div className="d-inline-block dropdown">
							{/* <DropDownButton
								showArrowIcon={false}
								stylingMode="text"
								icon="bx bx-bell bx-tada"
								items={pendingItems}
								itemTemplate="notificationItem"
							>
								<Template name="notificationItem" render={renderItems} />
							</DropDownButton> */}
						</div>
					</div>
				</div>
			</header>
		</React.Fragment>
	)
}

export default Header
