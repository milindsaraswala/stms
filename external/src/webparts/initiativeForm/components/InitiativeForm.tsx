import { WebPartContext } from "@microsoft/sp-webpart-base"
import { sp } from "@pnp/sp"
import { IItemAddResult, IItemUpdateResult } from "@pnp/sp/items"
import "@pnp/sp/items/index"
import "@pnp/sp/lists"
import { PermissionKind } from "@pnp/sp/security"
import "@pnp/sp/webs"
import { AxiosResponse } from "axios"
import Form, {
	AsyncRule,
	ButtonItem,
	ButtonOptions,
	GroupItem,
	Label,
	RangeRule,
	RequiredRule,
	SimpleItem,
} from "devextreme-react/form"
import { LoadPanel } from "devextreme-react/load-panel"
import "devextreme-react/number-box"
import "devextreme-react/text-area"
import { custom } from "devextreme/ui/dialog"
import React, { useEffect, useRef, useState } from "react"
import { approveRejectSelectedTask, getWFData } from "../../../services"
import { IInitiativeFormProps } from "./IInitiativeFormProps"

export const InitiativeFormWebPartContext = React.createContext<WebPartContext>(null)

export const InitiativeForm: React.FunctionComponent<IInitiativeFormProps> = (props: IInitiativeFormProps) => {
	const [readOnly, setReadOnly] = useState(true)
	const [defaultFormData, setDefaultFormData] = useState({})
	const [id, setId] = useState(0)
	const [action, setAction] = useState("")
	const [addButton, setAddButton] = useState(false)
	const [updateButton, setUpdateButton] = useState(false)
	const [deleteButton, setDeleteButton] = useState(false)
	const [approveRejectButton, setApproveRejectButton] = useState(false)
	const [types, setTypes] = useState([])
	const [divisions, setDivisions] = useState([])
	const [departments, setDepartments] = useState([])
	const [executionDepartments, setExecutionDepartments] = useState([])
	const [statuses, setStatuses] = useState([])
	const [health, setHealth] = useState([])
	const [selectedDivisionId, setSelectedDivisionId] = useState(0)
	const [plannedStartDate, setPlannedStartDate] = useState(new Date())
	const [actualStartDate, setActualStartDate] = useState(new Date())
	const [btnCount, setBtnCount] = useState(1)
	const [loadPanelVisible, setLoadPanelVisible] = useState(false)
	const initiativeFormRef = useRef() as React.MutableRefObject<Form>

	const position = { of: "#divInitiative" }

	useEffect(() => {
		sp.setup({
			spfxContext: props.context,
		})
		setId(+localStorage.getItem("initiativeID"))
		setAction(localStorage.getItem("initiativeAction"))
		setLoadPanelVisible(true)
	}, [])

	useEffect(() => {
		getListItemById()
		getTypes()
		getDivisions()
		getDepartments(false)
		getStatues()
		getHealth()
		if (action !== "") {
			showHideBtnAddNew()
			visibleApproveButton()
			if (action === "Add") setLoadPanelVisible(false)
		}
	}, [id, action])

	useEffect(() => {
		hasPermisssion()
		if (Object.keys(defaultFormData).length > 0) setLoadPanelVisible(false)
	}, [defaultFormData])

	useEffect(() => {
		getDepartments(true)
	}, [selectedDivisionId])

	const addlog = (id: number, stage: string, pendingWithId: number): Promise<IItemAddResult> => {
		const addNewData = initiativeFormRef.current.instance.option("formData")
		const list = sp.web.lists.getByTitle("InitiativesAuditTrail")
		return new Promise((resolve) => {
			resolve(
				list.items.add({
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
					Stage: stage,
					PendingWithId: pendingWithId,
					DivisionsId: addNewData["DivisionsId"],
					Planned: addNewData["Planned"],
					RefNumber: id,
					ActionById: props.context.pageContext.legacyPageContext.userId,
				})
			)
		})
	}

	const getListItemById = async () => {
		if (id === 0) return
		const items = await sp.web.lists.getByTitle("Initiatives").items.getById(id).get()
		setDefaultFormData(items)
	}

	const visibleApproveButton = () => {
		const absoluteUrl = props.context.pageContext.web.absoluteUrl
		const listName = "Initiatives"
		const getRunningWFTaskForCurrentUserListItemSOAPUrl = `${absoluteUrl}/_vti_bin/NintexWorkflow/Workflow.asmx?op=GetRunningWorkflowTasksForCurrentUserForListItem`
		const getRunningWFTaskForCurrentUserListItemSOAPRequest = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
																																<soap:Body>
																																	<GetRunningWorkflowTasksForCurrentUserForListItem xmlns="http://nintex.com">
																																		<itemId>${id}</itemId>
																																		<listName>${listName}</listName>
																																	</GetRunningWorkflowTasksForCurrentUserForListItem>
																																</soap:Body>
																															 </soap:Envelope>`
		const header = { "Content-Type": "text/xml" }

		if (action !== "Add") {
			getWFData(
				getRunningWFTaskForCurrentUserListItemSOAPUrl,
				getRunningWFTaskForCurrentUserListItemSOAPRequest,
				header
			).then((getRunningWFTaskForCurrentUserListItemResponse: AxiosResponse<any>) => {
				const parser = require("fast-xml-parser")

				if (parser.validate(getRunningWFTaskForCurrentUserListItemResponse.data) === true) {
					const getRunningWFTaskForCurrentUserListItem = parser.parse(
						getRunningWFTaskForCurrentUserListItemResponse.data
					)
					const getRunningWFTaskForCurrentUserListItemResult =
						getRunningWFTaskForCurrentUserListItem["soap:Envelope"]["soap:Body"]
							.GetRunningWorkflowTasksForCurrentUserForListItemResponse
							.GetRunningWorkflowTasksForCurrentUserForListItemResult
					if (
						getRunningWFTaskForCurrentUserListItemResult.UserTask !== undefined &&
						defaultFormData["Stage"] === "Pending" &&
						defaultFormData["PendingWithId"] === props.context.pageContext.legacyPageContext.userId
					) {
						setBtnCount((btnCount) => btnCount + 3)
						setApproveRejectButton(true)
					}
				}
			})
		}
	}

	const updateInitiative = (e: any) => {
		const validationResult = e.validationGroup.validate()
		if (validationResult.isValid) {
			const updatedFormData = initiativeFormRef.current.instance.option("formData")
			const list = sp.web.lists.getByTitle("Initiatives")
			list.items
				.getById(id)
				.update({
					Title: updatedFormData["Title"],
					TypesId: updatedFormData["TypesId"],
					DepartmentId: updatedFormData["DepartmentId"],
					Completion: updatedFormData["Completion"],
					PlannedStartDate: updatedFormData["PlannedStartDate"],
					PlannedFinishDate: updatedFormData["PlannedFinishDate"],
					ActualStartDate: updatedFormData["ActualStartDate"],
					ActualFinishDate: updatedFormData["ActualFinishDate"],
					KeyAchievements: updatedFormData["KeyAchievements"],
					KeyBottlenecks: updatedFormData["KeyBottlenecks"],
					StatusId: updatedFormData["StatusId"],
					HealthId: updatedFormData["HealthId"],
					ExecutionDepartmentId: updatedFormData["ExecutionDepartmentId"],
					ProjectID: updatedFormData["ProjectID"],
					DANumber: updatedFormData["DANumber"],
					Comments: updatedFormData["Comments"],
					DivisionsId: updatedFormData["DivisionsId"],
					Planned: updatedFormData["Planned"],
				})
				.then((iar) => {
					const absoluteUrl = props.context.pageContext.web.absoluteUrl
					approveRejectSelectedTask(absoluteUrl, id, "Initiatives", "Update").then((d) => {
						addlog(id, "Update", 0).then((logResult) => {
							document.location.href = "/"
						})
					})
				})
				.catch((error: any) => {
					console.log("Error: ", error)
				})
		}
	}

	const addInitiative = (e) => {
		const validationResult = e.validationGroup.validate()
		if (validationResult.isValid) {
			const addNewData = initiativeFormRef.current.instance.option("formData")
			const list = sp.web.lists.getByTitle("Initiatives")
			list.items
				.add({
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
					DivisionsId: addNewData["DivisionsId"],
					Planned: addNewData["Planned"],
				})
				.then((iar: IItemAddResult) => {
					addlog(iar.data.Id, "Pending", iar.data.PendingWithId).then((logResult) => {
						document.location.href = "/"
					})
				})
				.catch((error: any) => {
					console.log("Error: ", error)
				})
		}
	}

	const deleteInitiative = () => {
		const myDialog = custom({
			showTitle: false,
			messageHtml: "<b><i>Are you sure you want to delete this record??</i></b>",
			buttons: [
				{
					text: "Yes",
					onClick: (): boolean => {
						return true
					},
				},
				{
					text: "No",
					onClick: (): boolean => {
						return false
					},
				},
			],
		})
		myDialog.show().then((dialogResult: boolean) => {
			if (dialogResult) {
				sp.web.lists
					.getByTitle("Initiatives")
					.items.getById(id)
					.update({
						Stage: "Deleted",
					})
					.then((iur: IItemUpdateResult) => {
						addlog(id, "Deleted", 0).then((logResult) => {
							document.location.href = "/"
						})
					})
			}
		})
	}

	const cancelForm = (e) => {
		location.href = "/"
	}

	const showHideBtnAddNew = async () => {
		const list = sp.web.lists.getByTitle("Initiatives")
		const perms = await list.effectiveBasePermissions()

		if (list.hasPermissions(perms, PermissionKind.AddListItems) && action === "Add") {
			setBtnCount((btnCount) => btnCount + 1)
			setAddButton(true)
			setReadOnly(false)
		}
	}

	const hasPermisssion = async () => {
		const list = sp.web.lists.getByTitle("Initiatives")
		const perms = await list.effectiveBasePermissions()

		if (
			list.hasPermissions(perms, PermissionKind.EditListItems) &&
			action === "Update" &&
			((defaultFormData["PendingWithId"] === props.context.pageContext.legacyPageContext.userId &&
				defaultFormData["Stage"] === "Amended") ||
				(defaultFormData["AuthorId"] === props.context.pageContext.legacyPageContext.userId &&
					defaultFormData["Stage"] === "Approved"))
		) {
			setBtnCount((btnCount) => btnCount + 1)
			setUpdateButton(true)
			setReadOnly(false)
		}

		if (
			list.hasPermissions(perms, PermissionKind.DeleteListItems) &&
			action === "Update" &&
			defaultFormData["AuthorId"] === props.context.pageContext.legacyPageContext.userId &&
			(defaultFormData["PendingWithId"] === props.context.pageContext.legacyPageContext.userId ||
				defaultFormData["Stage"] === "Pending")
		) {
			setBtnCount((btnCount) => btnCount + 1)
			setDeleteButton(true)
		}

		if (
			defaultFormData["PendingWithId"] === props.context.pageContext.legacyPageContext.userId &&
			defaultFormData["Stage"] === "Pending"
		) {
			setBtnCount((btnCount) => btnCount + 3)
			setApproveRejectButton(true)
		}
	}

	const approveInitiative = () => {
		const absoluteUrl = props.context.pageContext.web.absoluteUrl
		approveRejectSelectedTask(absoluteUrl, id, "Initiatives", "Approve").then((d) => {
			addlog(id, "Approved", 0).then((logResult) => {
				document.location.href = "/"
			})
		})
	}

	const rejectInitiative = () => {
		const absoluteUrl = props.context.pageContext.web.absoluteUrl
		approveRejectSelectedTask(absoluteUrl, id, "Initiatives", "Reject").then((d) => {
			addlog(id, "Rejected", 0).then((logResult) => {
				document.location.href = "/"
			})
		})
	}

	const amendInitiative = () => {
		const absoluteUrl = props.context.pageContext.web.absoluteUrl
		approveRejectSelectedTask(absoluteUrl, id, "Initiatives", "Amend").then((d) => {
			addlog(id, "Amended", 0).then((logResult) => {
				document.location.href = "/"
			})
		})
	}

	const getTypes = async () => {
		const items = await sp.web.lists
			.getByTitle("InitiativeTypes")
			.items.select("Id", "Title")
			.filter("Status eq 'Active'")
			.getAll()
		setTypes(items)
	}

	const getDivisions = async () => {
		const items = await sp.web.lists
			.getByTitle("Divisions")
			.items.select("Id", "Title")
			.filter("Status eq 'Active'")
			.getAll()
		setDivisions(items)
	}

	const getDepartments = async (filterWithDivisionId: boolean) => {
		const filterString = filterWithDivisionId ? `and DivisionId eq ${selectedDivisionId}` : ""
		const items = await sp.web.lists
			.getByTitle("Departments")
			.items.select("Id", "Title")
			.filter(`Status eq 'Active' ${filterString}`)
			.getAll()
		filterWithDivisionId ? setDepartments(items) : setExecutionDepartments(items)
	}

	const getStatues = async () => {
		const items = await sp.web.lists
			.getByTitle("Status")
			.items.select("Id", "Title")
			.filter("Status eq 'Active'")
			.getAll()
		setStatuses(items)
	}

	const getHealth = async () => {
		const items = await sp.web.lists
			.getByTitle("Health")
			.items.select("Id", "Title")
			.filter("Status eq 'Active'")
			.getAll()
		setHealth(items)
	}

	const selectOptions = (dataSource: any): object => {
		return {
			dataSource: dataSource,
			valueExpr: "Id",
			displayExpr: "Title",
		}
	}

	const getDivisionSelectionChanged = (e) => {
		setSelectedDivisionId(e.value)
	}

	const divisionEditorOptions = {
		dataSource: divisions,
		onValueChanged: getDivisionSelectionChanged,
		valueExpr: "Id",
		displayExpr: "Title",
	}

	const onValueChangedPlannedStartDate = (e) => {
		setPlannedStartDate(e.value)
	}

	const onValueChangedActualStartDate = (e) => {
		setActualStartDate(e.value)
	}

	const customizeItem = (e) => {
		if (e.itemType === "simple") {
			if (!e.editorOptions) e.editorOptions = {}
			e.editorOptions.validationMessageMode = "always"
		}
	}

	const asyncTestValidation = (params: any): Promise<boolean> => {
		return new Promise((resolve) => {
			sp.web.lists
				.getByTitle("Initiatives")
				.items.filter(`Title eq '${params.value}'`)
				.getAll()
				.then((items) => {
					resolve(action === "Add" && items.length > 0)
				})
		})
	}

	return (
		<InitiativeFormWebPartContext.Provider value={props.context}>
			<LoadPanel
				shadingColor="rgba(0,0,0,0.4)"
				position={position}
				visible={loadPanelVisible}
				showIndicator={true}
				shading={true}
				showPane={true}
				closeOnOutsideClick={false}
			/>
			<div id="#divInitiative">
				<Form
					formData={defaultFormData}
					labelLocation="top"
					colCount={2}
					ref={initiativeFormRef}
					readOnly={readOnly}
					validationGroup="initiatorValidatorGroup"
					customizeItem={customizeItem}
				>
					<SimpleItem dataField="TypesId" editorType="dxSelectBox" editorOptions={selectOptions(types)}>
						<Label text="Initiative Type" />
						<RequiredRule message="Initiative type is required." />
					</SimpleItem>
					<SimpleItem dataField="DivisionsId" editorType="dxSelectBox" editorOptions={divisionEditorOptions}>
						<Label text="Division" />
						<RequiredRule message="Division is required." />
					</SimpleItem>
					<SimpleItem dataField="DepartmentId" editorType="dxSelectBox" editorOptions={selectOptions(departments)}>
						<Label text="Business Department" />
						<RequiredRule message="Department is required." />
					</SimpleItem>
					<SimpleItem
						dataField="Completion"
						editorType="dxNumberBox"
						editorOptions={{ format: "#0%", step: 0.25, min: 0, max: 1 }}
					>
						<RequiredRule message="Completion is required." />
					</SimpleItem>
					<SimpleItem dataField="Title" colSpan={2}>
						<Label text="Initiative Name" />
						<RequiredRule message="Initiative Name is required." />
						<AsyncRule message="Initiative Name already exists." validationCallback={asyncTestValidation} />
					</SimpleItem>
					<SimpleItem
						dataField="PlannedStartDate"
						editorType="dxDateBox"
						editorOptions={{ displayFormat: "dd-MMM-yyyy", onValueChanged: onValueChangedPlannedStartDate }}
					/>
					<SimpleItem
						dataField="PlannedFinishDate"
						editorType="dxDateBox"
						editorOptions={{ displayFormat: "dd-MMM-yyyy", min: plannedStartDate }}
					>
						<RangeRule min={plannedStartDate} message="Planned finish date must be greater than planned start date." />
					</SimpleItem>
					<SimpleItem
						dataField="ActualStartDate"
						editorType="dxDateBox"
						editorOptions={{ displayFormat: "dd-MMM-yyyy", onValueChanged: onValueChangedActualStartDate }}
					/>
					<SimpleItem
						dataField="ActualFinishDate"
						editorType="dxDateBox"
						editorOptions={{ displayFormat: "dd-MMM-yyyy", min: actualStartDate }}
					>
						<RangeRule min={actualStartDate} message="Actual finish date must be greater than actual start date." />
					</SimpleItem>
					<SimpleItem dataField="KeyAchievements" colSpan={2} editorType="dxTextArea" editorOptions={{ height: 90 }} />
					<SimpleItem dataField="KeyBottlenecks" colSpan={2} editorType="dxTextArea" editorOptions={{ height: 90 }} />
					<SimpleItem dataField="StatusId" editorType="dxSelectBox" editorOptions={selectOptions(statuses)}>
						<Label text="Status" />
						<RequiredRule message="Status is required" />
					</SimpleItem>
					<SimpleItem dataField="HealthId" editorType="dxSelectBox" editorOptions={selectOptions(health)}>
						<Label text="Health" />
						<RequiredRule message="Health is required" />
					</SimpleItem>
					<SimpleItem
						dataField="ExecutionDepartmentId"
						editorType="dxSelectBox"
						editorOptions={selectOptions(executionDepartments)}
					>
						<Label text="Execution Department" />
					</SimpleItem>
					<SimpleItem dataField="Planned" editorType="dxCheckBox">
						<Label text="Planned (Y/N)" />
					</SimpleItem>
					<SimpleItem dataField="ProjectID">
						<Label text="Project ID (for projects)" />
					</SimpleItem>
					<SimpleItem dataField="DANumber">
						<Label text="DA Number" />
					</SimpleItem>
					<SimpleItem dataField="Comments" colSpan={2} editorType="dxTextArea" editorOptions={{ height: 90 }} />
					<GroupItem colCount={btnCount} cssClass="buttons" colSpan="2">
						<ButtonItem cssClass="buttons-column" visible={addButton}>
							<ButtonOptions
								text="Save"
								type="default"
								stylingMode="contained"
								width={210}
								height={40}
								useSubmitBehavior={true}
								onClick={addInitiative}
							/>
						</ButtonItem>
						<ButtonItem cssClass="buttons-column" visible={approveRejectButton}>
							<ButtonOptions
								text="Approve"
								type="default"
								stylingMode="contained"
								width={210}
								height={40}
								onClick={approveInitiative}
							/>
						</ButtonItem>
						<ButtonItem cssClass="buttons-column" visible={approveRejectButton}>
							<ButtonOptions
								text="Reject"
								type="default"
								stylingMode="contained"
								width={210}
								height={40}
								onClick={rejectInitiative}
							/>
						</ButtonItem>
						<ButtonItem cssClass="buttons-column" visible={approveRejectButton}>
							<ButtonOptions
								text="Amend"
								type="default"
								stylingMode="contained"
								width={210}
								height={40}
								onClick={amendInitiative}
							/>
						</ButtonItem>
						<ButtonItem cssClass="buttons-column" visible={updateButton}>
							<ButtonOptions
								text="Update"
								type="default"
								stylingMode="contained"
								width={210}
								height={40}
								onClick={updateInitiative}
							/>
						</ButtonItem>
						<ButtonItem cssClass="buttons-column" visible={deleteButton}>
							<ButtonOptions
								text="Delete"
								type="default"
								stylingMode="contained"
								width={210}
								height={40}
								onClick={deleteInitiative}
							/>
						</ButtonItem>
						<ButtonItem cssClass="buttons-column">
							<ButtonOptions
								text="Cancel"
								type="default"
								stylingMode="contained"
								width={210}
								height={40}
								onClick={cancelForm}
							/>
						</ButtonItem>
					</GroupItem>
				</Form>
			</div>
		</InitiativeFormWebPartContext.Provider>
	)
}
