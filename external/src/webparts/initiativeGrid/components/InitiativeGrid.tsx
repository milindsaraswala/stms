import { WebPartContext } from "@microsoft/sp-webpart-base"
import { sp } from "@pnp/sp"
import "@pnp/sp/items"
import "@pnp/sp/lists"
import { PermissionKind } from "@pnp/sp/security"
import "@pnp/sp/webs"
import DataGrid, {
	Button,
	Column,
	ColumnFixing,
	Editing,
	Export,
	FilterRow,
	Grouping,
	GroupItem,
	GroupPanel,
	HeaderFilter,
	Pager,
	Paging,
	Scrolling,
	SearchPanel,
	Selection,
	Summary,
	TotalItem,
} from "devextreme-react/data-grid"
import "devextreme-react/drop-down-button"
import { exportDataGrid } from "devextreme/excel_exporter"
import cBox from "devextreme/ui/check_box"
import { custom } from "devextreme/ui/dialog"
import ExcelJS from "exceljs"
import saveAs from "file-saver"
// import { jsPDF } from "jspdf"
// import "jspdf-autotable"
import React, { useEffect, useRef, useState } from "react"
import Completion from "../../../components/Completion"
import { approveRejectSelectedTask } from "../../../services"
import { IInitiativeGridProps } from "./IInitiativeGridProps"

export const InitiativeGridWebPartContext = React.createContext<WebPartContext>(null)

export const InitiativeGrid: React.FunctionComponent<IInitiativeGridProps> = (props: IInitiativeGridProps) => {
	const pageSizes = [10, 25, 50, 100]
	const [gridData, setGridData] = useState([])
	const [disabledKeys, setDisabledKeys] = useState([])
	const [collapsed, setCollasped] = useState(false)
	const gridRef = useRef() as React.MutableRefObject<DataGrid>

	useEffect(() => {
		sp.setup({
			spfxContext: props.context,
		})
		getListItems()
		gridRef.current.instance.beginCustomLoading("loading ...")
	}, [])

	useEffect(() => {
		if (gridData.length >= 0) gridRef.current.instance.endCustomLoading()
	}, [gridData])

	const getListItems = async () => {
		const allItems: any[] = await sp.web.lists
			.getByTitle("Initiatives")
			.items.select(
				"ID",
				"Title",
				"TypesId",
				"Types/Title",
				"Types/TColor",
				"DivisionsId",
				"Divisions/Title",
				"DepartmentId",
				"Department/Title",
				"Completion",
				"Author/Title",
				"PlannedStartDate",
				"PlannedFinishDate",
				"ActualStartDate",
				"ActualFinishDate",
				"StatusId",
				"Status/Title",
				"HealthId",
				"Health/Title",
				"Health/HColor",
				"ExecutionDepartmentId",
				"ExecutionDepartment/Title",
				"Planned",
				"DANumber",
				"PendingWithId",
				"Stage"
			)
			.expand("Types", "Divisions", "Author", "Department", "Status", "Health", "ExecutionDepartment")
			.filter(
				`Stage ne 'Deleted' and (Stage eq 'Approved' or PendingWith/Title eq '${props.context.pageContext.user.displayName}' or Author/Title eq '${props.context.pageContext.user.displayName}')`
			)
			.getAll()
		setGridData(allItems)
	}

	const onContentReady = (e) => {
		if (!collapsed) {
			e.component.expandRow(["EnviroCare"])
			setCollasped(true)
		}
	}

	// const exportGrid = () => {
	// 	const doc = new jsPDF()
	// 	const dataGrid = gridRef.current.instance

	// 	exportDataGridToPdf({
	// 		jsPDFDocument: doc,
	// 		component: dataGrid,
	// 	}).then(() => {
	// 		doc.save("Initiatives.pdf")
	// 	})
	// }

	const onExporting = (e) => {
		const workbook = new ExcelJS.Workbook()
		const worksheet = workbook.addWorksheet("Initiatives")

		exportDataGrid({
			component: e.component,
			worksheet: worksheet,
			autoFilterEnabled: true,
			customizeCell: ({ gridCell, excelCell }) => {
				if (gridCell.rowType === "data") {
					if (gridCell.column.dataField === "Planned") {
						excelCell.value = gridCell.value ? "Yes" : "No"
					}
				}
			},
		}).then(() => {
			workbook.xlsx.writeBuffer().then((buffer) => {
				saveAs(new Blob([buffer], { type: "application/octet-stream" }), "Initiative.xlsx")
			})
		})
		e.cancel = true
	}

	const onToolbarPreparing = (e) => {
		const toolbarItems = e.toolbarOptions.items

		toolbarItems.splice(
			0,
			0,
			{
				widget: "dxButton",
				location: "after",
				options: {
					text: "Approve Selected",
					type: "default",
					stylingMode: "contained",
					width: 150,
					onClick: () => {
						const ids: number[] = gridRef.current.instance.getSelectedRowKeys()
						approveSelected(ids)
					},
				},
			},
			{
				widget: "dxButton",
				location: "after",
				options: {
					text: "Add New",
					type: "default",
					stylingMode: "contained",
					width: 120,
					onClick: () => {
						localStorage.setItem("initiativeAction", "Add")
						localStorage.setItem("initiativeID", "0")
						location.href = "/SitePages/IntiativeForm.aspx"
					},
				},
			}
			// {
			// 	widget: "dxButton",
			// 	location: "after",
			// 	options: {
			// 		text: "Export to PDF",
			// 		type: "default",
			// 		stylingMode: "contained",
			// 		onClick: exportGrid,
			// 	},
			// }
		)
	}

	const updateRow = (e: any) => {
		localStorage.setItem("initiativeID", e.row.key)
		localStorage.setItem("initiativeAction", "Update")
		location.href = "/SitePages/IntiativeForm.aspx"
	}

	const deleteRow = (e: any) => {
		const myDialog = custom({
			showTitle: false,
			messageHtml: "<b><i>Are you sure you want to delete this record??</i></b>",
			buttons: [
				{
					text: "Yes",
					onClick: (): boolean => {
						return true
					},
					stylingMode: "contained",
					type: "default",
				},
				{
					text: "No",
					onClick: (): boolean => {
						return false
					},
					stylingMode: "contained",
					type: "default",
				},
			],
		})
		myDialog.show().then((dialogResult: boolean) => {
			if (dialogResult) {
				sp.web.lists
					.getByTitle("Initiatives")
					.items.getById(e.row.key)
					.update({
						Stage: "Deleted",
					})
					.then(() => {
						const addNewData = e.row.data
						sp.web.lists
							.getByTitle("InitiativesAuditTrail")
							.items.add({
								Title: addNewData["Title"],
								TypesId: addNewData["TypesId"],
								DepartmentId: addNewData["DepartmentId"],
								Completion: addNewData["Completion"],
								PlannedStartDate: addNewData["PlannedStartDate"],
								PlannedFinishDate: addNewData["PlannedFinishDate"],
								ActualStartDate: addNewData["ActualStartDate"],
								ActualFinishDate: addNewData["ActualFinishDate"],
								KeyAchievements: addNewData["KeyAchievements"],
								KeyBottlenecks: addNewData["KeyBottlenecks"],
								StatusId: addNewData["StatusId"],
								HealthId: addNewData["HealthId"],
								ExecutionDepartmentId: addNewData["ExecutionDepartmentId"],
								ProjectID: addNewData["ProjectID"],
								DANumber: addNewData["DANumber"],
								Comments: addNewData["Comments"],
								Stage: "Deleted",
								PendingWithId: addNewData["PendingWithId"],
								DivisionsId: addNewData["DivisionsId"],
								Planned: addNewData["Planned"],
								RefNumber: e.row.key,
								ActionById: props.context.pageContext.legacyPageContext.userId,
							})
							.then(() => {
								gridRef.current.instance.deleteRow(e.row.rowIndex)
							})
					})
			}
		})
	}

	const onCellPrepared = (e: any) => {
		const currentUserId = props.context.pageContext.legacyPageContext.userId
		if (
			e.rowType === "data" &&
			e.column.command === "select" &&
			(e.data.PendingWithId !== currentUserId || e.data.Stage !== "Pending")
		) {
			const cbElement = e.cellElement.getElementsByClassName("dx-select-checkbox")
			const cbInstance = cBox.getInstance(cbElement[0])
			cbInstance.option("visible", false)
			if (disabledKeys.indexOf(e.data.ID) === -1) {
				setDisabledKeys([...disabledKeys, e.data.ID])
			}
		}

		if (e.rowType === "data" && e.column.dataField === "Health.Title") {
			e.cellElement.style.backgroundColor = e.data.Health.HColor
		}

		if (e.rowType === "data" && e.column.dataField === "Types.Title") {
			e.cellElement.style.backgroundColor = e.data.Types.TColor
		}
	}

	const onSelectionChanged = (e: any) => {
		if (disabledKeys.length > 0) e.component.deselectRows(disabledKeys)
	}

	const approveSelected = (ids: number[]) => {
		const rowIndexes = []
		gridRef.current.instance.beginCustomLoading("Selected initiative are getting approve")
		const promises = ids.map((id) => {
			const absoluteUrl = props.context.pageContext.web.absoluteUrl
			const listName = "Initiatives"
			approveRejectSelectedTask(absoluteUrl, id, listName, "Approve")
			rowIndexes.push(gridRef.current.instance.getRowIndexByKey(id))
		})
		Promise.all(promises).then((results) => {
			rowIndexes.map((rowIndex) => {
				const cbElement = gridRef.current.instance
					.getCellElement(rowIndex, 0)
					.getElementsByClassName("dx-select-checkbox")
				const cbInstance = cBox.getInstance(cbElement[0])
				cbInstance.option("visible", false)
			})

			setTimeout(() => {
				gridRef.current.instance.endCustomLoading()
			}, 5000)
		})
	}

	const hasPermission = async (): Promise<boolean> => {
		const list = sp.web.lists.getByTitle("Initiatives")
		try {
			const perms = await list.effectiveBasePermissions()
			return list.hasPermissions(perms, PermissionKind.DeleteListItems)
		} catch (e) {
			return false
		}
	}

	const allowDeleting = (e): boolean => {
		return (
			hasPermission() &&
			e.row.data.Author.Title === props.context.pageContext.user.displayName &&
			(e.row.data.PendingWithId === props.context.pageContext.legacyPageContext.userId ||
				e.row.data.Stage === "Pending")
		)
	}

	return (
		<InitiativeGridWebPartContext.Provider value={props.context}>
			<DataGrid
				dataSource={gridData}
				showBorders={true}
				onContentReady={(e) => onContentReady(e)}
				keyExpr="ID"
				onToolbarPreparing={onToolbarPreparing}
				ref={gridRef}
				onCellPrepared={onCellPrepared}
				onSelectionChanged={onSelectionChanged}
				rowAlternationEnabled={true}
				columnAutoWidth={true}
				onExporting={onExporting}
			>
				<Editing confirmDelete={false} allowUpdating={true} allowDeleting={allowDeleting} useIcons={true} />
				<Selection mode="multiple" selectAllMode="allPages" showCheckBoxesMode="always" />
				<FilterRow visible={true} applyFilter={true} />
				<GroupPanel visible={true} />
				<SearchPanel visible={true} highlightCaseSensitive={true} />
				<Grouping autoExpandAll={false} />
				<HeaderFilter visible={true} />
				<Summary>
					<GroupItem column="ID" summaryType="count" />
					<TotalItem
						column="ID"
						summaryType="count"
						displayFormat="Total records are {0}"
						alignment="left"
						showInColumn="Title"
					/>
				</Summary>

				<Column dataField="Title" caption="Initiative Name" dataType="string" fixed={true} />
				<Column dataField="Types.Title" caption="Initiative Types" dataType="string" fixed={true} />
				<Column dataField="Divisions.Title" caption="Division" dataType="string" fixed={true} />
				<Column dataField="Department.Title" caption="Department" dataType="string" fixed={true} />
				<Column
					dataField="Completion"
					dataType="number"
					cellRender={Completion}
					format="percent"
					alignment="left"
					allowGrouping={false}
					cssClass="bullet"
					fixed={true}
				/>
				<Column dataField="Author.Title" caption="Initiator Name" dataType="string" />
				<Column dataField="PlannedStartDate" dataType="date" format="dd-MMM-yyyy" />
				<Column dataField="PlannedFinishDate" dataType="date" format="dd-MMM-yyyy" />
				<Column dataField="ActualStartDate" dataType="date" format="dd-MMM-yyyy" />
				<Column dataField="ActualFinishDate" dataType="date" format="dd-MMM-yyyy" />
				<Column dataField="Status.Title" caption="Status" />
				<Column dataField="Health.Title" caption="Health" />
				<Column dataField="ExecutionDepartment.Title" caption="Execution Department" />
				<Column dataField="Planned" />
				<Column dataField="ProjectID" caption="Project ID" />
				<Column dataField="DANumber" caption="DA Number" />
				<Column dataField="Stage" caption="Stage" />

				<Column type="buttons" caption="Action">
					<Button name="edit" onClick={updateRow} hint="View/Edit" />
					<Button name="delete" onClick={deleteRow} hint="Delete" />
				</Column>

				<Export enabled={true} allowExportSelectedData={true} />
				<Scrolling mode="standard" />
				<ColumnFixing enabled={true} />
				<Pager allowedPageSizes={pageSizes} showPageSizeSelector={true} showInfo={true} />
				<Paging defaultPageSize={10} />
			</DataGrid>
		</InitiativeGridWebPartContext.Provider>
	)
}
