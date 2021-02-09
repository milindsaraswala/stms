import axios, { AxiosResponse } from "axios"

export const dateFormat = (date: Date, fstr: string, utc: any) => {
	utc = utc ? "getUTC" : "get"
	return fstr.replace(/%[YmdHMS]/g, (m) => {
		switch (m) {
			case "%Y":
				return date[utc + "FullYear"]() // no leading zeros required
			case "%m":
				m = 1 + date[utc + "Month"]()
				break
			case "%d":
				m = date[utc + "Date"]()
				break
			case "%H":
				m = date[utc + "Hours"]()
				break
			case "%M":
				m = date[utc + "Minutes"]()
				break
			case "%S":
				m = date[utc + "Seconds"]()
				break
			default:
				return m.slice(1) // unknown code, remove %
		}
		// add leading zero if required
		return ("0" + m).slice(-2)
	})
}

export const getWFData = (soapURL: string, soapRequest: string, header: {}): Promise<AxiosResponse<any>> => {
	return axios.post(soapURL, soapRequest, { headers: header })
}

export const approveRejectSelectedTask = (
	absoluteUrl: string,
	id: number,
	listName: string,
	outcome: string
): Promise<string> => {
	return new Promise((resolve) => {
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

		getWFData(
			getRunningWFTaskForCurrentUserListItemSOAPUrl,
			getRunningWFTaskForCurrentUserListItemSOAPRequest,
			header
		).then((getRunningWFTaskForCurrentUserListItemResponse: AxiosResponse<any>) => {
			const parser = require("fast-xml-parser")

			if (parser.validate(getRunningWFTaskForCurrentUserListItemResponse.data) === true) {
				const getRunningWFTaskForCurrentUserListItem = parser.parse(getRunningWFTaskForCurrentUserListItemResponse.data)
				const getRunningWFTaskForCurrentUserListItemResult =
					getRunningWFTaskForCurrentUserListItem["soap:Envelope"]["soap:Body"]
						.GetRunningWorkflowTasksForCurrentUserForListItemResponse
						.GetRunningWorkflowTasksForCurrentUserForListItemResult
				if (getRunningWFTaskForCurrentUserListItemResult.UserTask !== undefined) {
					const spTaskId = getRunningWFTaskForCurrentUserListItemResult.UserTask.SharePointTaskId
					const comments = ""
					const taskListName = "Workflow Tasks"

					const processFlexiTaskResponse2SOAPUrl = `${absoluteUrl}/_vti_bin/NintexWorkflow/Workflow.asmx?op=ProcessFlexiTaskResponse2`
					const processFlexiTaskResponse2SOAPRequest = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
																												<soap:Body>
																													<ProcessFlexiTaskResponse2 xmlns="http://nintex.com">
																														<comments>${comments}</comments>
																														<outcome>${outcome}</outcome>
																														<spTaskId>${spTaskId}</spTaskId>
																														<taskListName>${taskListName}</taskListName>
																													</ProcessFlexiTaskResponse2>
																												</soap:Body>
																											</soap:Envelope>`

					getWFData(processFlexiTaskResponse2SOAPUrl, processFlexiTaskResponse2SOAPRequest, header).then(
						(processFlexiTaskResponse2Response: AxiosResponse<any>) => {
							if (parser.validate(processFlexiTaskResponse2Response.data) === true) {
								const processFlexiTaskResponse2 = parser.parse(processFlexiTaskResponse2Response.data)
								const processFlexiTaskResponse2Result =
									processFlexiTaskResponse2["soap:Envelope"]["soap:Body"].ProcessFlexiTaskResponse2Response
										.ProcessFlexiTaskResponse2Result

								resolve(processFlexiTaskResponse2)
							}
						}
					)
				} else {
					resolve("")
				}
			}
		})
	})
}
