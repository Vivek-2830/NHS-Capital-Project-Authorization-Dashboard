import * as React from 'react';
import styles from './CaptialForm.module.scss';
import { ICaptialFormProps } from './ICaptialFormProps';
import { escape, times } from '@microsoft/sp-lodash-subset';
import { Item, sp, SPSharedObjectType } from '@pnp/sp/presets/all';
import * as moment from 'moment';
import {  ChoiceGroup, DatePicker, 
          DefaultButton, 
          DetailsList, 
          Dialog, DialogType, Dropdown, 
          IColumn, 
          Icon, 
          IIconProps, 
          Label, 
          Link, 
          PrimaryButton, 
          SearchBox, 
          TextField, 
          Toggle,
          } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker"; 
import { Web } from '@pnp/sp/webs';
import { toRelativeUrl } from '@pnp/spfx-controls-react';


export interface ICaptialFormState {
  CaptialFormData : any;
  AddCapitalFormDialog : boolean;
  Title: any;
  IsUrgent: boolean;
  PurposeID: any;
  Purpose: any;
  ProjectName: any;
  ProjectDescription: any;
  Location: any;
  SpecificLocation: any;
  CapitalPlanCategory: any;
  FundingSource: any;
  Approvalamount: any;
  ProjectSponsorID: any;
  ProjectSponsorTitle: any;
  ContactEmail: any;
  ApprovedBudget: any;
  EstimatedProjectCompletion: any;
  CostCentre: any;
  IsEstatesImplications: boolean;
  EstatesImplications: any;
  ImplicationIT: boolean;
  Implications: any;
  Status: any;
  ReadReviewerComment: any;
  ApproverComment: any;
  ProjectManagerID: any;
  CapitalPurposelist: any;
  CapitalLocationlist: any;
  CapitalPlanCategorylist: any;
  CapitalCostCenter: any;
  CapitalStatuslist: any;
  ProjectSponsor: any;
  ProjectManager: any;
  CapitalFundingSourcelist: any;
  searchText: any;
  CapitalExportData: any;
  startDate: any;
  endDate: any;
  EditIsUrgent: any;
  EditPurpose: any;
  EditProjectName: any;
  EditProjectDescription: any;
  EditLocation: any;
  EditSpecificLocation: any;
  EditCapitalPlanCategory: any;
  EditFundingSource: any;
  EditApprovalamount: any;
  EditProjectSponsor: any;
  EditProjectManagerID: any;
  EditProjectManager: any;
  EditContactEmail: any;
  EditApprovedBudget: any;
  EditEstimatedProjectCompletion: any;
  EditCostCentre: any;
  EditIsEstatesImplications: any;
  EditEstatesImplications: any;
  EditStatus: any;
  EditImplicationIT: any;
  EditImplications: any;
  EditFilterDialog: boolean;
  CurrentCapitalFormID: any;
  UpdateCapitalFormFilterDialog: any;
  EditReviewerComment: any;
  DeleteCurrentitem: any;
  DeleteFilterDialog:  boolean;
  SearchData : any;
  ReadFilterDialog: boolean;
  ReadAllData: any;
  ReadIsUrgent: any;
  ReadPurpose: any;
  ReadProjectName: any;
  ReadProjectDescription: any;
  ReadLocation: any;
  ReadSpecificLocation: any;
  ReadCapitalPlanCategory: any;
  ReadFundingSource: any;
  ReadProjectSponsor: any;
  ReadProjectManager: any;
  ReadContactEmail: any;
  ReadApprovedBudget: any;
  ReadEstimatedProjectCompletion: any;
  ReadCostCentre: any;
  ReadIsEstatesImplications: any;
  ReadEstatesImplications: any;
  ReadStatus: any;
  ReadImplicationIT: any;
  ReadImplications: any;
  ReadApprovalamount: any;
  ReadProjectManagerID: any;
  ApprovalUser: boolean;
  ReviewerUser: boolean;
  ReturnFilter: boolean;
  ApproveFilter: boolean;
  IsReviewer: boolean;
  IsApproval: boolean;
  ReviewerComment: any;
  ReadReviewerID: any;
  ReviewerFilterDialog: boolean;
  ApprovalComment: any;
  ReadApprovalID: any;
  ReadApproverComment:  any;
  CapitalFilterExportData : any;
  AllStatus: any;
  Pending: any;
  Approved: any;
  Rejected : any;
  ExportStatus : any;
  AllDocuments: any;
  GetAllDocuments: any;
  TempId:number;
  IncremenetState:number;
  DeleteDocuments:any;
  EditRequestId: any;
  CurrentDocumentID: any;
  ReadRequestId: any;
  DeleteDocument : any;
}

const addIcon: IIconProps = { iconName: 'Add' };

const AddCapitalFormDialogContentProps = {
  title: "Add Captialize Form"
};

const ReadCapitalFormDialogContentProps = {
  title: "Read Capitalize Form"
};

const UpdateCapitalFormDialogContentProps = {
  title: "Update Captialize Form"
};

const DeleteFilterDialogContentProps = {
};

const addmodelProps = {
  className: "Add-Dialog"
};

const readmodelProps = {
  className: "Read-Dialog"
};

const updatemodelProps = {
  className : "Update-Dialog"
};

const deletmodelProps = {
  className : "Delete-Form"
};

const SendIcon : IIconProps = { iconName: 'Send'};

const CancelIcon : IIconProps = { iconName: 'Cancel'};

const TextDocumentEdit : IIconProps = { iconName: 'TextDocumentEdit' };

require("../assets/css/style.css");
require("../assets/css/fabric.min.css");


export default class CaptialForm extends React.Component<ICaptialFormProps, ICaptialFormState> {

  constructor(props: ICaptialFormProps, state: ICaptialFormState){
    super(props);

    this.state = {
      CaptialFormData : "",
      AddCapitalFormDialog: true,
      Title: "",
      IsUrgent: true,
      PurposeID: "",
      Purpose:"",
      ProjectName: "",
      ProjectDescription: "",
      Location: "",
      SpecificLocation: "",
      CapitalPlanCategory: "",
      FundingSource: "",
      Approvalamount: "",
      ProjectSponsorID: "",
      ProjectSponsorTitle: "",
      ContactEmail: "",
      ApprovedBudget: "",
      EstimatedProjectCompletion: "",
      CostCentre: "",
      IsEstatesImplications: true,
      EstatesImplications: "",
      ImplicationIT: true,
      Implications:"",
      Status: "",
      ReadReviewerComment: "",
      ApproverComment: "",
      ProjectManagerID: [],
      CapitalPurposelist: "",
      CapitalLocationlist: "",
      CapitalPlanCategorylist: "",
      CapitalFundingSourcelist: "",
      CapitalCostCenter: "",
      CapitalStatuslist: "",
      ProjectSponsor: "",
      ProjectManager: "",
      searchText: "",
      CapitalExportData: "",
      startDate : "",
      endDate:  "",
      EditIsUrgent :"",
      EditPurpose: "",
      EditProjectName: "",
      EditProjectDescription: "",
      EditLocation: "",
      EditSpecificLocation: "",
      EditCapitalPlanCategory: "",
      EditFundingSource: "",
      EditApprovalamount: "",
      EditProjectSponsor: "",
      EditProjectManagerID: "",
      EditProjectManager: "",
      EditContactEmail: "",
      EditApprovedBudget: "",
      EditEstimatedProjectCompletion: "",
      EditCostCentre: "",
      EditIsEstatesImplications: "",
      EditEstatesImplications: "",
      EditStatus: "",
      EditImplicationIT: "",
      EditImplications: "",
      EditFilterDialog : true,
      CurrentCapitalFormID :"",
      UpdateCapitalFormFilterDialog: "",
      EditReviewerComment: "",
      DeleteCurrentitem: "",
      DeleteFilterDialog :true,
      ReadFilterDialog : true,
      SearchData: "",
      ReadAllData: "",
      ReadIsUrgent: "",
      ReadPurpose:  "",
      ReadProjectName: "",
      ReadProjectDescription: "",
      ReadLocation: "",
      ReadSpecificLocation: "",
      ReadCapitalPlanCategory: "",
      ReadFundingSource: "",
      ReadProjectSponsor: "",
      ReadProjectManager: "",
      ReadContactEmail: "",
      ReadApprovedBudget: "", 
      ReadEstimatedProjectCompletion: "",
      ReadCostCentre: "",
      ReadIsEstatesImplications: "",
      ReadEstatesImplications: "",
      ReadStatus:"",
      ReadImplicationIT: "",
      ReadImplications: "",
      ReadApprovalamount: "",
      ReadProjectManagerID: "",
      ApprovalUser: false,
      ReviewerUser: false,
      ReturnFilter: true,
      ApproveFilter: true,
      IsReviewer: false,
      IsApproval: false ,
      ReviewerComment: "",
      ReadReviewerID: "",
      ReviewerFilterDialog: false,
      ApprovalComment: "",
      ReadApprovalID: "",
      ReadApproverComment: "",
      CapitalFilterExportData: "",
      AllStatus: "",
      Pending : "",
      Approved : "",
      Rejected : "",
      ExportStatus : "",
      AllDocuments: [],
      GetAllDocuments: [],
      TempId:0,
      IncremenetState:1,
      DeleteDocuments:[],
      EditRequestId : "",
      CurrentDocumentID: "",
      ReadRequestId : "",
      DeleteDocument : []
    };

  }


  public render(): React.ReactElement<ICaptialFormProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    
  const columns: IColumn[] = [
    {
      key: "ID",
      name: "ID",
      fieldName: "ID",
      minWidth: 50,
      maxWidth: 50,
      isResizable: false
    },
    {
      key: "ProjectName",
      name: "Project Name",
      fieldName: "ProjectName",
      minWidth: 100,
      maxWidth: 200,
      isResizable: false
    },
    {
      key: "ProjectManager",
      name: "Project Manager",
      fieldName: "ProjectManager",
      minWidth: 100,
      maxWidth: 150,
      isResizable: false,
    },
    { 
      key: 'Status', 
      name: 'Status', 
      fieldName: 'Status', 
      minWidth: 100,
      maxWidth: 150, 
      isResizable: false, 
      onRender: (item) => {
          if(item.Status == "P1 Pending"){
            return   <div className='P1Pending'>{item.Status}</div>;
          }
          else if(item.Status == "P1 Returned"){
            return   <div className='P1Returned'>{item.Status}</div>;
          }
          else if(item.Status == "P1 Approved"){
            return <div className='P1Approved'>{item.Status}</div>;
          }
          else if(item.Status == "P2 Approved"){
            return <div className='P2Approved'>{item.Status}</div>;
          }
          else if(item.Status == "P2 Rejected"){
            return <div className='P2Rejected'>{item.Status}</div>;
          }
          else if(item.Status == "Approved"){
            return <div className='Approved'>{item.Status}</div>;
          }
          else if(item.Status == "Rejected"){
            return <div className='Rejected'>{item.Status}</div>;
          }
          else if(item.Status == "Packed"){
            return <div className='Packed'>{item.Status}</div>;
          }
      }
    },
    { 
      key: 'Modified', 
      name: 'Date', 
      fieldName: 'Modified', 
      minWidth: 100, 
      maxWidth: 100, 
      isResizable: false,
      onRender: (item) => {
        return <span>{moment(new Date(item.Modified)).format("DD-MM-YYYY")}</span>;
        // dddd [at] h:mm A
      }
    },
    { 
      key: 'Purpose', 
      name: 'Purpose', 
      fieldName: 'Purpose', 
      minWidth: 100, 
      maxWidth: 150, 
      isResizable: false 
    },
    {
      key: 'ReviewerComment', 
      name: 'Comments', 
      fieldName: 'ReviewerComment', 
      minWidth: 100, 
      maxWidth: 150, 
      isResizable: false 
    },
    {
      key: "Actions",
      name: "",
      fieldName: "",
      minWidth: 100,
      maxWidth: 250,
      onRender: (item) => (
        <>
        <div>
            <div className='ms-Grid-row'>
              <div className='ms-Grid-col Read-Icon'>
                <a style={{ fontSize: "20px" , cursor: "pointer", marginRight: "5px" ,  color: "black" }}>
                  <Icon iconName='RedEye12' onClick={() => this.setState({ GetAllDocuments:[], AllDocuments:[], ReadFilterDialog : false , CurrentCapitalFormID : item.ID }, () => {this.GetReadCapitalForm(item.ID) , this.GetDocumemtFromdocumentLibrary(item.ID);})}></Icon>
                </a>
              </div>
              {
                (item.Status == "P1 Returned" || item.Status == "P1 Pending") && (this.state.IsReviewer == false && this.state.IsApproval == false) ?
                <>
                  <div className='ms-Grid-col Edit-Icon'>
                    <a style={{ fontSize: "20px" , cursor: "pointer" , color: "green" }}>
                      <Icon className='Edit-Icon' iconName='Edit' onClick={() => this.setState({GetAllDocuments:[], AllDocuments:[] ,EditFilterDialog : false , CurrentCapitalFormID : item.ID ,CurrentDocumentID: item.ID }, () => {this.GetEditCapitalForm(item.ID), this.GetDocumemtFromdocumentLibrary(item.ID);})} />
                    </a>
                  </div>

                  
                </> 
                :<></>
              }
              {
                item.Status == "P1 Pending" && (this.state.IsReviewer == false && this.state.IsApproval == false) ? 
                <>
                  <div className='ms-Grid-col Delete-Icon'>
                    <a style={{ fontSize: "20px" , cursor: "pointer", color: "#ee3535" }}>
                      <Icon className='Delete-Icon' iconName='Delete' onClick={() => this.setState({ DeleteFilterDialog : false , DeleteCurrentitem :item.ID })} />
                    </a>
                  </div>
                </> 
                :
                <></>
              }

            </div>
          </div>
        </>
      )
    }
  ];

    return (
      
        <div className="captialForm">
          <div className='ms-Grid'>
            <div className='ms-Grid-row'>
              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>  
                <div className="d-flex-header">
                    {/* <div className="nhs-text">NHS</div> */}
                    <h3>Capital Project Authorization</h3>
                </div>
              </div>
            </div>

            <div className='ms-Grid-row'>
              <div className='fieldGroup'>
                  <div className="ms-Grid-col ms-sm9 ms-md9 ms-lg4">
                    <SearchBox placeholder="Search" className='new-Search new-Search-Animation' 
                      onChange={(e) => { this.runexportfunction(e.target.value) }}
                      onClear={(e) => { this.runexportfunction("")} }
                      />
                </div>
                <div className='ms-Grid-col ms-sm9 ms-sm9 ms-lg4'>
                    <DatePicker allowTextInput={false}  value={this.state.startDate ? this.state.startDate : ""} 
                      onSelectDate={(e) => { this.setState({ startDate: e },() => this.Export_onFilter() )}}
                      placeholder='Select Start Date..!!' aria-label='Select Start Date' />
                </div>
                <div className='ms-Grid-col ms-sm9 ms-sm9 ms-lg4'>
                <DatePicker allowTextInput={false}  value={this.state.endDate ? this.state.endDate : ""} 
                      onSelectDate={(e) => { this.setState({ endDate: e },() => this.Export_onFilter() )}}
                      placeholder='Select End Date..!!' aria-label='Select End Date' />
                </div>
                <div className='ms-Grid-col ms-sm3 ms-md8 ms-lg12'>
                  <div className='Add-Capital-Project'>
                    <PrimaryButton className='Add-Form' iconProps={addIcon} type='Add' text='Add' onClick={() => this.setState({ AddCapitalFormDialog : false , GetAllDocuments:[], AllDocuments:[] })}/>
                  </div>
                </div>
              </div>
            </div>
          </div>      
              
          <div className='ms-Grid-row'>
            <div className='ms-Grid-col'>
              <div className='Status-button'>
                    <div className='All-button'>
                        <PrimaryButton
                          text="All"
                          onClick={() => this.applyfilter("allStatus")}
                        />
                    </div>
                    <div className='Pending-button'>
                        <PrimaryButton
                          text="Pending"
                          onClick={() => this.applyfilter("pendingStatus")}
                        />
                    </div>
                    <div className='Approved-button'>
                        <PrimaryButton  
                          text="Approved"
                          onClick={() => this.applyfilter("approvedStatus")}
                        />
                    </div>
                    <div className='Rejected-button'>
                        <PrimaryButton
                          text="Rejected"
                          onClick={() => this.applyfilter("rejectStatus")}
                        />
                    </div>
              </div>
            </div>
          </div>
           
                {/* Add Capital Authorization Form */}
                <Dialog
                  hidden={this.state.AddCapitalFormDialog}
                  onDismiss={() => 
                    this.setState({
                      AddCapitalFormDialog : true,
                      IsUrgent: true,
                      Purpose: "",
                      ProjectName: "",
                      ProjectDescription: "",
                      Location: "",
                      SpecificLocation: "",
                      CapitalPlanCategory : "",
                      FundingSource: "",
                      Approvalamount: "",
                      ProjectSponsor: "",
                      ProjectManager: "",
                      ContactEmail: "",
                      ApprovedBudget: "",
                      EstimatedProjectCompletion: "",
                      CostCentre: "",
                      IsEstatesImplications: true,
                      EstatesImplications: "",
                      Implications: "",
                      ImplicationIT: true,
                      Status: "",
                      ReadReviewerComment: "",
                      ApproverComment: "",
                    })
                  }
                  dialogContentProps={AddCapitalFormDialogContentProps}
                  modalProps={addmodelProps}
                  minWidth={800}
                >
                  <div>
                    <div className='ms-Grid-row ms-md'>
                      <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        
                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-Isurgent'>
                            <Toggle
                              label="Is the Project genuinely urgent: *"
                              checked={this.state.IsUrgent}
                              onText='On'
                              offText='Off'
                              onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this.setState({ IsUrgent : checked})}
                            ></Toggle>
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-Purpose'>
                            <Dropdown
                              options={(this.state.CapitalPurposelist)}
                              label="Purpose of this C1 form:"
                              placeholder='Select C1 form'
                              required
                              onChange={(e, option, text) => 
                                this.setState({ Purpose: option.text})
                              }
                            />
                          </div>
                        </div>  

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-ProjectName'>
                            <TextField
                              label="Name of Projet(while for uplift include project number):"
                              name="ProjectName"
                              type="Text"
                              placeholder='Please Enter Your Name'
                              required={true}
                              onChange={(value) => 
                                this.setState({ ProjectName: value.target["value"]})
                              }
                              value={this.state.ProjectName}
                            />
                          </div>
                        </div>
                          
                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className="Add-ProjectDescription">
                            <TextField 
                              label="Detailed description of the Project/Asset/Equipment(s):"
                              className='Add-Text'
                              name="ProjectDescription"
                              type="Text"
                              multiline rows = {2}
                              required={true}
                              onChange={(value) =>
                                this.setState({ ProjectDescription: value.target["value"]})
                              }
                              value={this.state.ProjectDescription}
                            />
                          </div>
                        </div>
                         
                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-Location'>
                            <Dropdown
                              options={(this.state.CapitalLocationlist)}
                              label="Location of the Project/Asset/Equipment(s):"
                              placeholder='Select Location'
                              required
                              onChange={(e, option, text) => 
                                this.setState({ Location: option.text})
                              }
                            />
                          </div>
                        </div>
                          
                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-SpecificLocation'>
                            <TextField
                              label="Specific Location/Building/Cost Centre(s) of the Asset:"
                              name="SpecificLocation"
                              type="Text"
                              multiline rows= {2}
                              required={true}
                              onChange={(value) => 
                                this.setState({ SpecificLocation: value.target["value"]})
                              }
                              value={this.state.SpecificLocation}
                            />
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-CapitalPlan'>
                            <Dropdown
                              options={this.state.CapitalPlanCategorylist}
                              label="Capital Plan Category:"
                              placeholder='Select Capital Plan Category'
                              required
                              onChange={(e, option, text) => 
                                this.setState({CapitalPlanCategory : option.text })
                              }
                            />
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-FundingSource'>
                            <Dropdown
                              options={this.state.CapitalFundingSourcelist}
                              label="Funding Source:"
                              placeholder='Select Funding Source'
                              required
                              onChange={(e, option,text) => 
                                this.setState({ FundingSource : option.text })
                              }
                            />
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-ApprovedBudget'>
                            <TextField
                              label="Approval amount (£) requested on this C1 form (excluding recoverable VAT)"
                              required={true}
                              type="Number"
                              onChange={(value) =>
                                this.setState({ ApprovedBudget : value.target["value"]})
                              }
                              value={this.state.ApprovedBudget}
                            />
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-ProjectSponsor'>  
                            <TextField
                              label="Project Sponsor:"
                              name="ProjectSponsor"
                              type="Text"
                              required={true}
                              onChange={(value) => 
                                this.setState({ ProjectSponsor : value.target["value"]})
                              }
                              value={this.state.ProjectSponsor}
                            />
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-ProjectManager'> 
                            <PeoplePicker
                              context={this.props.context}
                              titleText="Project Manager:"
                              personSelectionLimit={3}
                              // groupName={"Team Site Owners"} // Leave this blank in case you want to filter from all users
                              showtooltip={true}
                              required={true}
                              // defaultSelectedUsers={this.state.ProjectManager.Title}
                              onChange={this._getPeoplePickerItems}
                              principalTypes={[PrincipalType.User]}
                              resolveDelay={300} 
                              ensureUser={true}
                            />
                          </div>
                        </div>


                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-ContactEmail'> 
                            <TextField
                              label="Please provide your contact email address:"
                              name="ContactEmail"
                              type="Text"
                              placeholder='Please Enter Your Email Address'
                              required={true}
                              onChange={(value) => 
                                this.setState({ContactEmail: value.target["value"]})
                              }
                              value={this.state.ContactEmail}
                            />
                          </div>
                        </div>
                      
                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-Approvalamount'>       
                            <TextField
                              label="Approved Budget Amount/Balance(£)for the Project at FY Budget Setting:"
                              aria-label='If youre unsure, please consult your Management Accountant for guidance'
                              name="ApprovedBudget"
                              type='Number'
                              required={true}
                              onChange={(value) => 
                                this.setState({ Approvalamount: value.target["value"]})
                              }
                              value={this.state.Approvalamount}
                            />
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-EstimatedProjectCompletion'>  
                            <TextField 
                              label="Estimated Project Completion/Forecast Date:"
                              name="EstimatedProjectCompletion"
                              type="Text"
                              required={true}
                              onChange={(value) =>
                                this.setState({EstimatedProjectCompletion : value.target["value"] })
                              }
                              value={this.state.EstimatedProjectCompletion}
                            />
                          </div>
                        </div>
                        
                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                            <div className='Add-EstatesImplications'> 
                              <TextField
                                className="Add-Text"
                                label='what are those Implications?*'
                                name="Estates Implications"
                                multiline rows= {2}
                                type="Text"
                                required={true}
                                onChange={(value) => 
                                  this.setState({ EstatesImplications : value.target["value"] })
                                }
                                value={this.state.EstatesImplications}
                              />
                            </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                            <div className='Add-IsEstatesImplications'> 
                              <Toggle
                                className="Add-Text"
                                label="Are there any Estates implications?"
                                onText='On'
                                offText='Off'
                                checked={this.state.IsEstatesImplications}
                                onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this.setState({ IsEstatesImplications : checked})}
                              />
                            </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-CostCentre'> 
                            <Dropdown
                              className="Add-Text"
                              options={this.state.CapitalCostCenter}
                              label="Cost Centre :"
                              placeholder='Select Cost Centre'
                              required
                              onChange={(e, option, text) => 
                                this.setState({ CostCentre : option.text})
                              }
                            />
                          </div>
                        </div>
                        
                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-ImplicationIT'> 
                                <Toggle
                                  className="Add-Text"  
                                  label="Are there any IT implications?*"
                                  onText='On'
                                  offText='Off'
                                  checked={this.state.ImplicationIT}
                                  onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this.setState({ ImplicationIT : checked})}
                                />
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Statment'>
                            <Label className='Add-Statment'>Statement of need  : *</Label>
                          </div>
                              <label className='Attachmentlabel' htmlFor="Statement Document" >Choose files</label>
                                <input style={{display:'none'}} id="Statement Document" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "Statement Document","")}></input>
                                <div className='Document-wrapers'>
                                    {
                                        this.state.AllDocuments.length > 0 && (
                                          this.state.AllDocuments.map((item) => {
                                            return (
                                              <>
                                                {
                                                  item.text == "Statement Document" ?
                                                    <>
                                                      <div className='FilesWrap'>
                                                        <Label >{item.key.name}</Label>
                                                        <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.key.name)}/>
                                                      </div>
                                                    </> : <></>
                                                }
                                              </>
                                            );
                                          })
                                        )
                                    }
                              </div>  
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Add-Implications'> 
                                <TextField
                                  label="what are those Implications?"
                                  name="Implications"
                                  multiline rows= {2}
                                  type="Text"
                                  required={true}
                                  onChange={(value) =>
                                    this.setState({Implications : value.target["value"]})
                                  }
                                  value={this.state.Implications}
                                />
                            </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className="MedicalApproval">
                            <Label className="Add-MedicalApproval">Upload Medical Equipment Group Approval  : *</Label>
                          </div>
                              <label className='Attachmentlabel' htmlFor="Medical Approval Document" >Choose files</label>
                                <input style={{display:'none'}} id="Medical Approval Document" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "Medical Approval Document","")}></input>
                                  <div className='Document-wrapers'>
                                    {
                                        this.state.AllDocuments.length > 0 && (
                                          this.state.AllDocuments.map((item) => {
                                            return (
                                              <>
                                                {
                                                  item.text == "Medical Approval Document" ?
                                                    <>
                                                      <div className='FilesWrap'>
                                                        <Label >{item.key.name}</Label>
                                                        <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.key.name)}/>
                                                      </div>
                                                    </> : <></>
                                                }
                                              </>
                                            );
                                          })
                                        )
                                    }
                                </div>
                          
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='C1Form'>
                            <Label className='Add-C1Form'>C1 - Appendix(Breakdown of Capital Budget)<br /> and Upload the completed form.  :* </Label>
                          </div>
                              <label className='Attachmentlabel' htmlFor="C1 Form Document" >Choose files</label>
                                <input style={{display:'none'}} id="C1 Form Document" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "C1 Form Document","")}></input>
                                <div className='Document-wrapers'>
                                  {
                                      this.state.AllDocuments.length > 0 && (
                                        this.state.AllDocuments.map((item) => {
                                          return (
                                            <>
                                              {
                                                item.text == "C1 Form Document" ?
                                                  <>
                                                    <div className='FilesWrap'>
                                                      <Label >{item.key.name}</Label>
                                                        <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.key.name)}/>
                                                    </div>
                                                  </> : <></>
                                               }
                                            </>
                                          );
                                        })
                                      )
                                  }
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Otherdoc'>
                            <Label className='Add-Otherdoc'>Upload Other Documents : </Label>
                          </div>
                              <label className='Attachmentlabel' htmlFor="Other Docs" >Choose files</label>
                                <input style={{display:'none'}} id="Other Docs" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "Other Documents","")}></input>
                                <div className='Document-wrapers'>
                                  {
                                      this.state.AllDocuments.length > 0 && (
                                        this.state.AllDocuments.map((item) => {
                                          return (
                                            <>
                                              {
                                                item.text == "Other Documents" ?
                                                  <>
                                                    <div className='FilesWrap'>
                                                      <Label >{item.key.name}</Label>
                                                      <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.key.name)}/>
                                                    </div>
                                                  </> : <></>
                                              }
                                            </>
                                          );
                                        })
                                      )
                                  }
                                  </div>
                        </div>

                      </div>
                    </div>

                    <div className='ms-Grid-row'>
                        <div className='Submit-AuthorizationForm'>
                            <PrimaryButton
                              iconProps={SendIcon}
                              type="Submit"
                              text="Submit"
                              onClick={() => this.CapitalFormAdd()}
                            />

                            <DefaultButton
                              iconProps={CancelIcon}  
                              type="Cancel"
                              text="Cancel"
                              onClick={() => this.setState({ AddCapitalFormDialog : true })}
                            />
                        </div>
                    </div>
                 
                </div>
                </Dialog>

                {/* Read Capital Authorization Form */}
                <Dialog
                  hidden={this.state.ReadFilterDialog}
                  onDismiss={() =>
                    this.setState({
                      ReadFilterDialog : true,
                      IsUrgent: true,
                      Purpose: "",
                      ProjectName: "",
                      ProjectDescription: "",
                      Location: "",
                      SpecificLocation: "",
                      CapitalPlanCategory : "",
                      FundingSource: "",
                      Approvalamount: "",
                      ProjectSponsor: "",
                      ProjectManager: "",
                      ContactEmail: "",
                      ApprovedBudget: "",
                      EstimatedProjectCompletion: "",
                      CostCentre: "",
                      IsEstatesImplications: true,
                      EstatesImplications: "",
                      Implications: "",
                      ImplicationIT: true,
                      Status: "",
                      ReadReviewerComment: "",
                      ApproverComment: "",
                    })
                  }
                  dialogContentProps={ReadCapitalFormDialogContentProps}
                  modalProps={readmodelProps}
                  minWidth={800}
                >
                  {/* <Icon iconName='PageLeft'  onClick={() => this.setState({ ReadFilterDialog : true })}></Icon> */}
                  <div>
                    <div className='ms-Grid-row ms-md'>
                      <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                          <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                              <Label className='Read-Text'>Is the Project genuinely urgent: * <span>
                                <Icon iconName='Info' title='(e.g If there is an immediate danger to staff or patients, a delay in patient treatment, or any other urgent
                                concern, please document details in the statement of need)'></Icon>
                              </span></Label>
                              <p className='Read-p'>{this.state.IsUrgent == true ? "Yes" : "No"}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                            <Label className='Read-Text'>Purpose of this C1 form *</Label>
                             <p className='Read-p'>{this.state.ReadPurpose}</p>
                          </div>  

                          <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                            <Label className='Read-Text'>Name of Project (while for uplift include project number): * </Label>
                              <p className='Read-p'>{this.state.ReadProjectName}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                            <Label className='Read-Text'> Detailed description of the Project/Asset/Equipment(s): * </Label>
                              <p className='Read-p'>{this.state.ReadProjectDescription}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Location of the Project/Asset/Equipment(s): * </Label>
                              <p className='Read-p'>{this.state.ReadLocation}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Specific Location/Building/Cost Centre(s) of the Asset * <span><Icon iconName='Info' 
                            title='(where will the intended project/Asset/Equipment be stationed for verification/Inspection)
                              ** Equipment & IT include the Department and Cost Centre.
                              ** Site & Construction include the Building and Block Name'></Icon></span></Label>
                              <p className='Read-p'>{this.state.ReadSpecificLocation}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Capital Plan Category: * </Label>
                              <p className='Read-p'>{this.state.ReadCapitalPlanCategory}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Funding Source: * <span>
                              <Icon iconName='Info' title="(Select the appropriate project funding source. If you're unsure, please consult your Management
                                Accountant for guidance)"></Icon></span></Label>
                              <p className='Read-p'>{this.state.ReadFundingSource}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Approval amount (£) requested on this C1 form (excluding recoverable VAT)  * <span>
                              <Icon iconName='Info' title='For VAT queries, please get in touch with'></Icon>
                              </span></Label> 
                              <p className='Read-p'>{this.state.ReadApprovedBudget}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Project Sponsor:  * </Label>
                              <p className='Read-p'>{this.state.ReadProjectSponsor}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'>  Project Manager: * <span>
                              <Icon iconName='Info' title='(This individual is responsible for reviewing the requisition and invoice)'></Icon>
                              </span></Label>
                              <p className='Read-p'>{this.state.ReadProjectManager}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'>  Please provide your contact email address: * <span>
                              <Icon iconName='Info' title='(Enter an email address where we can reach you regarding any follow-ups, clarifications, or important
                                updates)'></Icon></span></Label>
                              <p className='Read-p'>{this.state.ReadContactEmail}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Approved Budget Amount/Balance(£) for the Project at FY Budget Setting * <span>
                              <Icon iconName='Info' title='(If youre unsure, please consult your Management Accountant for guidance)'>
                              </Icon></span></Label>
                              <p className='Read-p'>{this.state.ReadApprovalamount}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'>  Estimated Project Completion/Forecast Date * <span>
                              <Icon iconName="Info" title='Specify the phasing of the spending if it spans multiple financial years. Break it down by year, for example:
                                  Current year (CY) = 2024/25
                                  CY = £X, XXX
                                  CY+1 = £X, XXX
                                  CY+2 = £X, XXX, etc.'>
                                </Icon></span></Label>
                              <p className='Read-p'>{this.state.ReadEstimatedProjectCompletion}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Cost Centre : *  <span>
                              <Icon iconName='Info' title='(For any revenue/depreciation implications)'>
                                </Icon></span></Label>
                              <p className='Read-p'>{this.state.ReadCostCentre}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Are there any Estates implications? * <span>
                              <Icon iconName='Info' title="(E.g. additional power/AHU/reconfiguration of space)">
                                </Icon></span></Label>
                              <p className='Read-p'>{this.state.ReadIsEstatesImplications == true ? "Yes" : "No"}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> what are those Implications? * <span>
                              <Icon iconName='Info' title='(Ensure that all financial implications are fully accounted for in the total project or asset amount)'></Icon></span></Label>
                              <p className='Read-p'>{this.state.ReadEstatesImplications}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Status * </Label>
                              <p className='Read-p'>{this.state.ReadStatus}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'> Are there any IT implications? * <span>
                              <Icon iconName='Info' title="E.g. Additional server/ integration works/ hardware">
                                </Icon></span></Label>
                              <p className='Read-p'>{this.state.ReadImplicationIT == true ? "Yes" : "No"}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <Label className='Read-Text'>what are those Implications?  * <span>
                              <Icon iconName='Info' title='(Ensure that all financial implications are fully accounted for in the total project or asset amount)'>
                                </Icon></span></Label>
                              <p className='Read-p'>{this.state.ReadImplications}</p>
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                            <div className='Read-Statment'>
                              <Label className='Read-Text'>Statement of need  : *</Label>
                                  <label htmlFor="Statement Document"></label>
                                    <input style={{display:'none'}} id="Statement Document" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "Statement Document",this.state.CurrentCapitalFormID)}></input>
                                      {
                                          this.state.GetAllDocuments.length > 0 && (
                                            this.state.GetAllDocuments.map((item) => {
                                              return (
                                                <>
                                                  {
                                                    item.DocumentType == "Statement Document" && item.RequestId == this.state.CurrentCapitalFormID ?
                                                      <>
                                                        <div className='FilesWrap'>
                                                          <p>{item.Filename}</p>
                                                          {/* <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.key.name)}/> */}
                                                        </div>
                                                      </> : <></>
                                                  }
                                                </>
                                              );
                                            })
                                          )
                                      }
                            </div>
                          </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Read-Medical'>
                            <Label className='Read-Text'>Upload Medical Equipment Group Approval  : *</Label>
                              <label htmlFor="Medical Approval Document"></label>
                                <input style={{display:'none'}} id="Medical Approval Document" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "Medical Approval Document",this.state.CurrentCapitalFormID)}></input>
                                  {
                                    this.state.GetAllDocuments.length > 0 && (
                                      this.state.GetAllDocuments.map((item) => {
                                        return (
                                          <>
                                            {
                                              item.DocumentType == "Medical Approval Document" && item.RequestId == this.state.CurrentCapitalFormID ?
                                                <>
                                                  <div className='FilesWrap'>
                                                    <p>{item.Filename}</p>
                                                      {/* <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.key.name)}/> */}
                                                    </div>
                                                </> : <></>
                                            }
                                          </>
                                        );
                                      })
                                    )
                                  }
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                          <div className='Read-C1Form'>
                            <Label className='Read-Text'>C1 - Appendix(Breakdown of Capital Budget)<br /> and Upload the completed form.  :* </Label>
                              <label htmlFor="C1 Form Document"></label>
                                <input style={{display:'none'}} id="C1 Form Document" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "C1 Form Document",this.state.CurrentCapitalFormID)}></input>
                                  {
                                      this.state.GetAllDocuments.length > 0 && (
                                        this.state.GetAllDocuments.map((item) => {
                                          return (
                                            <>
                                              {
                                                item.DocumentType == "C1 Form Document" && item.RequestId == this.state.CurrentCapitalFormID ?
                                                  <>
                                                    <div className='FilesWrap'>
                                                      <p>{item.Filename}</p>
                                                      {/* <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.key.name)}/> */}
                                                    </div>
                                                  </> : <></>
                                              }
                                            </>
                                          );
                                        })
                                      )
                                  }
                          </div>
                        </div>

                        <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                          <div className='Read-OtherDoc'>
                            <Label className='Read-Text'>Upload Other Documents : </Label>
                              <label htmlFor="Other Docs"></label>
                                <input style={{display:'none'}} id="Other Docs" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "Other Documents",this.state.CurrentCapitalFormID)}></input>
                                  {
                                    this.state.GetAllDocuments.length > 0 && (
                                      this.state.GetAllDocuments.map((item) => {
                                          return (
                                              <>
                                                {
                                                  item.DocumentType == "Other Documents" && item.RequestId == this.state.CurrentCapitalFormID ?
                                                    <>
                                                      <div className='FilesWrap'>
                                                          <p>{item.Filename}</p>
                                                            {/* <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.Filename)}/> */}
                                                      </div>
                                                    </> : <></>
                                                }
                                              </>
                                            );
                                        })
                                      )
                                    }
                                    </div>
                        </div>


                          {
                            this.state.ReadStatus != "P1 Pending" ?
                            <>
                              <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                                <Label className='Read-Text'>Reviewer Comment * </Label>
                                <p className='Read-p'>{this.state.ReadReviewerComment}</p>
                              </div>
                            </>
                            :
                            <></>
                          }

                          {
                            this.state.ReadStatus == "P2 Approved" || this.state.ReadStatus == "P2 Rejected" ||  this.state.ReadStatus == "Approved"  || this.state.ReadStatus == "Rejected" ?
                            <>
                              <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                                <Label className='Read-Text'>Approval Comment * </Label>
                                <p className='Read-p'>{this.state.ReadApproverComment}</p>
                              </div>
                            </>
                            :
                            <></>
                          }

                          {
                            this.state.IsApproval == true &&  this.state.ReadStatus == "Packed" ?
                            <>
                              <div className="ms-Grid-col ms-sm12 ms-sm6 ms-lg6">
                                <div className='Read-Review-Comment'>
                                    <TextField
                                      label="Approver Comment *"
                                      name='Approver Comment'
                                      multiline rows= {2}
                                      type="Text"
                                      required={true}
                                      onChange={(value) => 
                                          this.setState({ ReadApproverComment : value.target["value"]})
                                      }
                                      value={this.state.ReadApproverComment}
                                    />
                                </div>
                              </div>


                                <div className='ms-Grid-row'>
                                  <div className="ms-Grid-col">
                                    <div className='Review-Comment'>
                                            
                                      <PrimaryButton
                                          text="Rejected"
                                          type="Rejected"
                                          onClick={() => this.setState({ ReadFilterDialog : true}, () => this.ApprovalControls("P2 Rejected" ,this.state.ReadApprovalID))}
                                      />

                                      <PrimaryButton
                                          text="Approved"
                                          onClick={() => this.setState({ ReadFilterDialog : true },() => this.ApprovalControls("P2 Approved" ,this.state.ReadApprovalID))}
                                      />

                                      <DefaultButton
                                          text="Cancel"
                                          type="Cancel"
                                          onClick={() => this.setState({ ReadFilterDialog : true })}
                                      />

                                    </div>
                                  </div>
                                </div>

                            </> 
                            : 
                            <></>
                          }
                          
                          {
                            this.state.IsReviewer == true && this.state.ReadStatus == "P1 Pending" ?  
                            <>
                                <div className="ms-Grid-col ms-sm12 ms-sm6 ms-lg6">
                                  <div className='Read-Review-Comment'>
                                    <TextField
                                      label="Reviewer Comment *"
                                      name='Reviewer Comment'
                                      multiline rows= {2}
                                      type="Text"
                                      required={true}
                                      onChange={(value) => 
                                          this.setState({ ReadReviewerComment : value.target["value"]})
                                        }
                                      value={this.state.ReadReviewerComment}
                                    />
                                  </div>
                                </div>

                                <div className='ms-Grid-row'>
                                  <div className="ms-Grid-col">
                                    <div className='Review-Comment'>
                                            
                                      <PrimaryButton
                                        text="Return"
                                        type="Return"
                                        onClick={() => this.setState({ ReadFilterDialog : true}, () => this.ReviewerControls("P1 Returned" ,this.state.ReadReviewerID))}
                                      />

                                      <PrimaryButton
                                        text="Approve"
                                        onClick={() => this.setState({ ReadFilterDialog : true },() => this.ReviewerControls("P1 Approved" , this.state.ReadReviewerID))}
                                      />

                                      <DefaultButton
                                        text="Cancel"
                                        type="Cancel"
                                        onClick={() => this.setState({ ReadFilterDialog : true })}
                                      />

                                    </div>
                                  </div>
                                </div>
                            </>
                            : 
                            <></>
                          }

                          {
                            this.state.IsApproval == true && this.state.ReadStatus == "P1 Approved" ? 
                            <>
                                <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                                    <div className='Read-Review-Comment'>
                                      <TextField
                                        label="ApprovalComment *"
                                        multiline rows= {2}
                                        type="Text"
                                        required={true}
                                        onChange={(value) => 
                                            this.setState({ ReadApproverComment : value.target["value"]})
                                        }
                                        value={this.state.ReadApproverComment}
                                      />
                                    </div>
                                </div>

                                        <div className='ms-Grid-row'>
                                          <div className='Approve-Comment'>
                                          <PrimaryButton
                                                    text="Approve"
                                                    onClick={() => this.setState({ ReadFilterDialog : true }, () => this.ApprovalControls("P2 Approved" ,this.state.ReadApprovalID))}
                                          />

                                          <PrimaryButton
                                                    text="Reject"
                                                    onClick={() => this.setState({ ReadFilterDialog : true }, () => this.ApprovalControls("P2 Rejected", this.state.ReadApprovalID))}
                                          />

                                          <PrimaryButton
                                                    text="Packed"
                                                    onClick={() => this.setState({ ReadFilterDialog : true }, () => this.ApprovalControls("Packed", this.state.ReadApprovalID))}
                                          />

                                          <DefaultButton
                                                    text="Cancel"
                                                    type="Cancel"
                                                    onClick={() => this.setState({ ReadFilterDialog : true })}
                                          />
                                        </div>
                                      </div>
                            </> 
                            : 
                            <></>
                          }

                          {
                            this.state.IsReviewer == true && (this.state.ReadStatus == "P2 Approved" || this.state.ReadStatus == "P2 Rejected") ?
                            <>
                                    <div className='ms-Grid-row'>
                                      <div className="Send-Summery">
                                        <PrimaryButton
                                          text="Send Summery"
                                          onClick={() => this.setState({ ReadFilterDialog : true }, () => this.SendSummeryControl(this.state.ReadApprovalID))}
                                        />

                                        <DefaultButton
                                            text="Cancel"
                                            type="Cancel"
                                            onClick={() => this.setState({ ReadFilterDialog : true })}
                                        />

                                      </div>
                                    </div> 
                            </>
                            :
                            <></>
                          }
 
                      </div>
                    </div>
                  </div>
                </Dialog>

                {/* Update Capital Authorization Form */}        
                <Dialog
                    hidden={this.state.EditFilterDialog}
                    onDismiss={() => 
                      this.setState({
                        EditFilterDialog : true,
                        IsUrgent: true,
                        Purpose: "",
                        ProjectName: "",
                        ProjectDescription: "",
                        Location: "",
                        SpecificLocation: "",
                        CapitalPlanCategory : "",
                        FundingSource: "",
                        Approvalamount: "",
                        ProjectSponsor: "",
                        ProjectManager: "",
                        ContactEmail: "",
                        ApprovedBudget: "",
                        EstimatedProjectCompletion: "",
                        CostCentre: "",
                        IsEstatesImplications: true,
                        EstatesImplications: "",
                        Implications: "",
                        ImplicationIT: true,
                        Status: "",
                        
                      })
                    }
                    dialogContentProps={UpdateCapitalFormDialogContentProps}
                    modalProps={updatemodelProps}
                    minWidth={1000}
                  >
                  
                    <div>
                      <div className='ms-Grid-row ms-md'>
                        <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>

                          <div className='Edit-IsUrgent'>                    
                              <Toggle
                                  label="Is the Project genuinely urgent: *"
                                  checked={this.state.IsUrgent}
                                  onText='On'
                                  offText='Off'
                                  onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this.setState({ IsUrgent : checked})}
                              />
                          </div>

                          <div className="Edit-Purpose">
                            <Dropdown
                                options={(this.state.CapitalPurposelist)}
                                label='Purpose of this C1 form:'
                                placeholder='Select C1 form'
                                required
                                defaultSelectedKey={this.state.EditPurpose}
                                onChange={(e, option, text) => 
                                  this.setState({ EditPurpose: option.text })
                                }
                            />
                          </div>

                          <div className='Edit-ProjectName'>
                              <TextField
                                label='Name of Projet(while for uplift include project number):'
                                name="ProjectName"
                                type="Text"
                                placeholder='Please Enter Your Name'
                                required={true}
                                onChange={(value) => 
                                  this.setState({ EditProjectName: value.target["value"]})
                                }
                                value={this.state.EditProjectName}
                              />
                          </div>

                          <div className='Edit-ProjectDescription'>
                              <TextField 
                                label=' Detailed description of the Project/Asset/Equipment(s):'
                                name="ProjectDescription"
                                type="Text"
                                multiline rows = {4}
                                required={true}
                                onChange={(value) =>
                                  this.setState({ EditProjectDescription: value.target["value"]})
                                }
                                value={this.state.EditProjectDescription}
                              />
                          </div>
                          
                          <div className="Edit-Location">
                              <Dropdown
                                options={(this.state.CapitalLocationlist)}
                                label='Location of the Project/Asset/Equipment(s):'
                                placeholder='Select Location'
                                required
                                defaultSelectedKey={this.state.EditLocation}
                                onChange={(e, option, text) => 
                                  this.setState({ EditLocation: option.text })
                                }
                              />
                          </div>

                          <div className='Edit-SpecificLocation'>
                              <TextField
                                label="Specific Location/Building/Cost Centre(s) of the Asset: "
                                name="SpecificLocation"
                                type="Text"
                                multiline rows= {4}
                                required={true}
                                onChange={(value) => 
                                  this.setState({ EditSpecificLocation: value.target["value"]})
                                }
                                value={this.state.EditSpecificLocation}
                              />
                          </div>

                          <div className='Edit-CapitalPlan'>
                              <Dropdown
                                options={this.state.CapitalPlanCategorylist}
                                label=" Capital Plan Category:"
                                placeholder='Select Capital Plan Category'
                                required
                                defaultSelectedKey={this.state.EditCapitalPlanCategory}
                                onChange={(e, option, text) => 
                                  this.setState({ EditCapitalPlanCategory : option.text })
                                }
                              />
                          </div>

                          <div className='Edit-ApproveBudget'>
                              <TextField
                                label=" Approval amount (£) requested on this C1 form (excluding recoverable VAT)  "
                                required={true}
                                type="Number"
                                onChange={(value) =>
                                  this.setState({ EditApprovedBudget : value.target["value"]})
                                }
                                value={this.state.EditApprovedBudget}
                              />
                          </div>

                          <div className='Edit-ProjectSponsor'>
                              <TextField
                                label=" Project Sponsor: "
                                name="ProjectSponsor"
                                type="Text"
                                required={true}
                                onChange={(value) => 
                                  this.setState({ EditProjectSponsor : value.target["value"]})
                                }
                                value={this.state.EditProjectSponsor}
                              />
                          </div>

                          <div className='Edit-ProjectManager'>
                              <PeoplePicker
                                context={this.props.context}
                                titleText="Project Manager"
                                personSelectionLimit={3}
                                // groupName={"Team Site Owners"} // Leave this blank in case you want to filter from all users
                                showtooltip={true}
                                required={true}
                                defaultSelectedUsers={[this.state.EditProjectManager]}
                                onChange={this._getPeoplePickerItems}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={300} 
                                ensureUser={true}
                              />
                          </div>

                          <div className='Edit-ContactEmail'>    
                              <TextField
                                label="Please provide your contact email address: "
                                name="ContactEmail"
                                type="Text"
                                placeholder='Please Enter Your Email Address'
                                required={true}
                                onChange={(value) => 
                                  this.setState({EditContactEmail: value.target["value"]})
                                }
                                value={this.state.EditContactEmail}
                              />
                          </div>

                          <div className='Edit-Approvalamount'>
                                <TextField
                                  label="Approved Budget Amount/Balance(£)for the Project at FY Budget Setting  "
                                  aria-label='If youre unsure, please consult your Management Accountant for guidance'
                                  name="ApprovedBudget"
                                  type='Number'
                                  required={true}
                                  onChange={(value) => 
                                    this.setState({ EditApprovalamount: value.target["value"]})
                                  }
                                  value={this.state.EditApprovalamount}
                                />
                          </div>

                          <div className='Edit-EstimatedProjectCompletion'>
                                <TextField 
                                  label="Estimated Project Completion/Forecast Date"
                                  name="EstimatedProjectCompletion"
                                  type="Text"
                                  required={true}
                                  onChange={(value) =>
                                    this.setState({EditEstimatedProjectCompletion : value.target["value"] })
                                  }
                                  value={this.state.EditEstimatedProjectCompletion}
                                />                            
                          </div>

                          <div className='Edit-CostCentre'>
                                <Dropdown
                                  options={this.state.CapitalCostCenter}
                                  label=" Cost Centre :  "
                                  placeholder='Select Cost Centre'
                                  required
                                  defaultSelectedKey={this.state.EditCostCentre}
                                  onChange={(e, option, text) => 
                                    this.setState({ EditCostCentre : option.text})
                                  }
                                />
                          </div>

                          <div className='Edit-IsEstatesImplications'>
                                  <Toggle
                                    label="Are there any Estates implications? "
                                    onText='On'
                                    offText='Off'
                                    checked={this.state.EditIsEstatesImplications}
                                    onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this.setState({ EditIsEstatesImplications : checked})}
                                  />
                          </div>

                          <div className="Edit-EstatesImplications">
                                  <TextField
                                    label='what are those Implications?*'
                                    name="Estates Implications"
                                    multiline rows= {4}
                                    type="Text"
                                    required={true}
                                    onChange={(value) => 
                                      this.setState({EditEstatesImplications : value.target["value"]})
                                    }
                                    value={this.state.EditEstatesImplications}
                                  />
                          </div>

                          <div className='Edit-ImplicationIT'>
                                  <Toggle
                                    label="Are there any IT implications? * "
                                    onText='On'
                                    offText='Off'
                                    checked={this.state.EditImplicationIT}
                                    onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => this.setState({ EditImplicationIT : checked})}
                                  />  
                          </div>
                                  
                          <div className='Edit-Implications'>
                                  <TextField
                                    label="what are those Implications?  "
                                    name="Implications"
                                    multiline rows= {4}
                                    type="Text"
                                    required={true}
                                    onChange={(value) =>
                                      this.setState({ EditImplications : value.target["value"]})
                                    }
                                    value={this.state.EditImplications}
                                  />
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <div className='Edit-Statement'>
                              <Label>Statement of need : * </Label>
                              </div>
                                <label className='Attachmentlabel' htmlFor="Statement Docs-Edit" >Choose files</label>
                                  <input style={{display:'none'}} id="Statement Docs-Edit" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "Statement Document",this.state.CurrentCapitalFormID)}></input>
                                    {
                                      this.state.GetAllDocuments.length > 0 && (
                                        this.state.GetAllDocuments.map((item) => {
                                          return (
                                              <>
                                                {
                                                  item.DocumentType == "Statement Document"  ?
                                                  // item.Documenttype == "Statement Document"  ?
                                                  // (item.Documenttype == "Statement Document" && item.RequestId == this.state.CurrentCapitalFormID) || (item.ID == "" && item.Documenttype == "Statement Document") ?
                                                  <>
                                                    <div className='FilesWrap'>
                                                      <Label>{item.Filename}</Label>
                                                          <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.Filename)}/>
                                                    </div>
                                                    </> : <></>
                                                }
                                              </>
                                            );
                                         })
                                        )
                                    }
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <div className='Edit-Medical'>
                              <Label>Upload Medical Equipment Group Approval : * </Label>
                            </div>
                                <label className='Attachmentlabel' htmlFor="Medical Docs-Edit" >Choose files</label>
                                  <input style={{display:'none'}} id="Medical Docs-Edit" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "Medical Approval Document",this.state.CurrentCapitalFormID)}></input>
                                    {
                                      this.state.GetAllDocuments.length > 0 && (
                                        this.state.GetAllDocuments.map((item) => {
                                          return (
                                            <>
                                              { 
                                                item.DocumentType == "Medical Approval Document" ?
                                                // (item.Documenttype == "Medical Approval Document" && item.RequestId == this.state.CurrentCapitalFormID) || (item.ID == "" && item.Documenttype == "Medical Approval Document") ?
                                                <>
                                                  <div className='FilesWrap'>
                                                     <Label>{item.Filename}</Label>
                                                        <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.Filename)}/>
                                                  </div>
                                                </> : <></>
                                              }
                                            </>
                                            );
                                          })
                                        )
                                    }
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <div className='Edit-C1Form'>
                              <Label>C1 - Appendix(Breakdown of Capital Budget)<br /> and Upload the completed form. :* </Label>
                            </div>
                                <label className='Attachmentlabel' htmlFor="C1 Form Document-Edit" >Choose files</label>
                                  <input style={{display:'none'}} id="C1 Form Document-Edit" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "C1 Form Document",this.state.CurrentCapitalFormID)}></input>
                                    {
                                      this.state.GetAllDocuments.length > 0 && (
                                        this.state.GetAllDocuments.map((item) => {
                                          return (
                                            <>
                                              {
                                                item.DocumentType == "C1 Form Document"  ?
                                                // (item.Documenttype == "C1 Form Document" && item.RequestId == this.state.CurrentCapitalFormID) || (item.ID == "" && item.Documenttype == "C1 Form Document") ?
                                                <>
                                                  <div className='FilesWrap'>
                                                    <Label>{item.Filename}</Label>
                                                      <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.Filename)}/>
                                                  </div>
                                                </> : <></>
                                              }
                                              </>
                                          );
                                        })
                                      )
                                    }
                          </div>

                          <div className='ms-Grid-col ms-sm12 ms-sm6 ms-lg6'>
                            <div className='Edit-Other'>
                              <Label>Upload Other Documents : </Label>
                            </div>
                                <label className='Attachmentlabel' htmlFor="Other Docs-Edit" >Choose files</label>
                                  <input style={{display:'none'}} id="Other Docs-Edit" type="file" multiple onChange={(e) => this.GetAttchments(e.target.files, "Other Documents",this.state.CurrentCapitalFormID)}></input>
                                    {
                                      this.state.GetAllDocuments.length > 0 && (
                                        this.state.GetAllDocuments.map((item) => {
                                          return (
                                            <>
                                              {
                                                item.DocumentType == "Other Documents"  ?
                                                // (item.Documenttype == "Other Documents" && item.RequestId == this.state.CurrentCapitalFormID) || (item.ID == "" && item.Documenttype == "Other Document") ?
                                                <>
                                                  <div className='FilesWrap'>
                                                    <Label>{item.Filename}</Label>
                                                      <Icon iconName="Cancel" onClick={(e)=>this.removeAtttchments(item.TempId,item.ID,item.Filename)}/>
                                                  </div>
                                                </> : <></>
                                              }
                                            </>
                                          );
                                        })
                                      )
                                    }
                          </div>

                        </div>
                      </div>
                      
                      <div className='ms-Grid-row'>
                          <div className='Update-AuthorizationForm'>
                              <PrimaryButton
                                iconProps={TextDocumentEdit}
                                type="Update"
                                text="Update"
                                onClick={() => this.UpdateCapitalForm(this.state.CurrentCapitalFormID)}
                              />

                              <DefaultButton
                                iconProps={CancelIcon}  
                                type="Cancel"
                                text="Cancel"
                                onClick={() => this.setState({ EditFilterDialog : true })}
                              />
                          </div>  
                      </div>
                    </div>
                </Dialog>
            
                {/* Delete Capital Authorization Form */}
                <Dialog
                      hidden={this.state.DeleteFilterDialog}
                      onDismiss={() => 
                        this.setState({
                          DeleteFilterDialog : true,
                        })
                      }
                      
                      dialogContentProps={DeleteFilterDialogContentProps}
                      modalProps={deletmodelProps}
                      minWidth={400}
                    >
                      <div className='Close-Icon'>
                        <Icon iconName='Cancel' onClick={() => this.setState({ DeleteFilterDialog : true })}></Icon>
                      </div>

                      <div className='Cancel-Icon'>
                          <Icon iconName='Cancel' className='cancel'/>
                      </div>

                      <div className='delete-Text'>
                          <h4>Are you sure?</h4>
                          <p>Do you really want to delete these record?</p>
                      </div>

                          <div className='ms-Grid-row'>
                            <div className='Delete-Form'>
                              <DefaultButton
                                type="Cancel"
                                text='Cancel'
                                onClick={() => this.setState({ DeleteFilterDialog : true })}
                              />

                              <PrimaryButton
                                type="Delete"
                                text="Delete"
                                onClick={() => this.DeleteCapitalForm()}
                              />
                            </div>
                          </div>  
                </Dialog>
                    
              <div className='ms-Grid'>
                <DetailsList
                  className='DetalisList-Form'
                  items={this.state.CaptialFormData}
                  columns={columns}
                  setKey="set"
                  layoutMode={1}
                  selectionMode={0}
                  isHeaderVisible={true}
                  ariaLabelForSelectionColumn="Toggle selection"
                  ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                  checkButtonAriaLabel="select row"
                >
                </DetailsList>
              </div>
        </div>
    );
  }

  public async componentDidMount() {
    this.GetCapitalForm();
    this.GetCaptialFormChoiceItems();
    this.GetCurrentUser();
    // this.GetDocumemtFromdocumentLibrary();
  }

/*Get User */
  public async GetCurrentUser() {
  let groups = await sp.web.currentUser.groups();
  console.log(groups);

  groups.forEach((items) => {
    if (items.Title == "Reviewer Group") {
      this.setState({ IsReviewer : true });
    } 
  });
  console.log(this.state.IsReviewer);

  groups.forEach((items) => {
    if(items.Title == "Approval Group") {
      this.setState({ IsApproval : true });
    }
  });
  console.log(this.state.IsApproval);

  // const currentuser = await sp.web.currentUser.get();
  // const group = await sp.web.siteGroups.getByName("Reviewer Group").users.get();
  // const userReviewerInGroup = group.some(user => user.Id == currentuser.Id);
  // const group2 = await sp.web.siteGroups.getByName("Approval Group").users.get();
  // const userApproveInGroup = group2.some(user => user.Id == currentuser.Id);

  // if(userReviewerInGroup) {
  //   this.setState({ ReviewerUser : true });
  //   console.log(this.state.ReviewerUser);
  //   console.log("Current User is in the Reviewer group.");

  // } else if(userApproveInGroup) {
  //   this.setState({ ApprovalUser : true });
  //   console.log(this.state.ApprovalUser);
  //   console.log("User is in the Approvers group but not the Reviewers group.");
  // }
}

/* Reviewer Control */
  public async ReviewerControls(Status,ID) {
  const updateform = await sp.web.lists.getByTitle("Capital Project Authorization").items.getById(ID).update({
      ReviewerComment: this.state.ReadReviewerComment,
      Status: Status,
    }).catch((error) => {
    console.log(error);
  });
  this.GetCapitalForm();
}

/* Approval Control */
  public async ApprovalControls(Status,ID) {
  const updateapproval = await sp.web.lists.getByTitle("Capital Project Authorization").items.getById(ID).update({
    ApproverComment: this.state.ReadApproverComment,
    Status: Status
  }).catch((error) => {
    console.log(error);
  });
  this.GetCapitalForm();
}

/* Send Summery */ 
  public async SendSummeryControl(ID) {
  const sendcontrol = await sp.web.lists.getByTitle("Capital Project Authorization").items.getById(ID).update({
    ReviewerComment: this.state.ReadReviewerComment,
    Status: this.state.ReadStatus == "P2 Approved" ? "Approved" : "Rejected"
  }).catch((error) => {
    console.log(error);
  });
  this.GetCapitalForm();
}

/* Get CapitalFormAdd Call */ 
  public async GetCapitalForm() {
    const captialformauth = await sp.web.lists.getByTitle("Capital Project Authorization").items.select(
      "ID",
      "Title",
      "IsUrgent",
      "Purpose/Title",
      "ProjectName",
      "ProjectDescription",
      "Location/Title",
      "SpecificLocation",
      "CapitalPlanCategory/Title",
      "FundingSource/Title",
      "Approvalamount",
      "ProjectSponsor",
      "ProjectManager/Title",
      "ProjectManager/ID",
      "ContactEmail",
      "ApprovedBudget",
      "EstimatedProjectCompletion",
      "CostCentre/Title",
      "IsEstatesImplications",
      "EstatesImplications",
      "ImplicationIT",
      "Implications",
      "Status/Title",
      "Status/ID",
      "ReviewerComment",
      "ApproverComment",
      "Author/Title",
      "Author/ID",
      "Modified",
      "Created"
      // "Purpose/ID",  
      // "Location/ID", 
      // "FundingSource/ID",
      // "CostCentre/ID",
      // "CapitalPlanCategory/ID",
    ).expand("ProjectManager","Author").get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(captialformauth);

      if(data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.Id ? item.Id : "",  
            IsUrgent: item.IsUrgent,
            Purpose: item.Purpose ? item.Purpose : "",
            ProjectName: item.ProjectName ? item.ProjectName : "",
            ProjectDescription: item.ProjectDescription ? item.ProjectDescription : "",
            Location: item.Location ? item.Location: "",
            SpecificLocation: item.SpecificLocation ? item.SpecificLocation : "",
            CapitalPlanCategory: item.CapitalPlanCategory ? item.CapitalPlanCategory : "",
            FundingSource: item.FundingSource? item.FundingSource: "",
            Approvalamount: item.Approvalamount ? item.Approvalamount : "",
            ProjectSponsor: item.ProjectSponsor ? item.ProjectSponsor : "",
            ProjectManager: item.ProjectManager.Title ? item.ProjectManager.Title : "",
            ProjectManagerID: item.ProjectManager.ID ? item.ProjectManager.ID : "",
            ContactEmail: item.ContactEmail ? item.ContactEmail : "",
            ApprovedBudget: item.ApprovedBudget ? item.ApprovedBudget : "",
            EstimatedProjectCompletion: item.EstimatedProjectCompletion ? item.EstimatedProjectCompletion : "",
            CostCentre: item.CostCentre ? item.CostCentre: "",
            IsEstatesImplications: item.IsEstatesImplications ,
            EstatesImplications: item.EstatesImplications ? item.EstatesImplications : "",
            ImplicationIT: item.ImplicationIT ,
            Status: item.Status ? item.Status : "",
            ReviewerComment: item.ReviewerComment ? item.ReviewerComment : "",
            ApproverComment: item.ApproverComment ? item.ApproverComment : "",
            Implications: item.Implications ? item.Implications : "",
            Modified: item.Modified ? item.Modified : "",
            Created: item.Created ? item.Created : ""
          });
        });
        this.setState({ CaptialFormData : AllData });
        this.setState({ CapitalExportData : AllData });
        console.log(this.state.CaptialFormData);
      }
    }).catch((Error) => {
      console.log("Error Retrived", Error);
    });
}

/* Add CapitalFormAdd Call */
  public async CapitalFormAdd() {
    if(this.state.IsUrgent == null || 
      this.state.Purpose.length == 0 || 
      this.state.ProjectName.length == 0 || 
      this.state.ProjectDescription.length == 0 ||
      this.state.Location.length == 0 ||
      this.state.SpecificLocation.length == 0 ||
      this.state.CapitalPlanCategory.length == 0 ||
      this.state.FundingSource.length == 0 ||
      this.state.Approvalamount.length == 0 ||
      this.state.ProjectSponsor.length == 0 ||
      this.state.ProjectManagerID.length == 0 ||
      this.state.ContactEmail.length == 0 ||
      this.state.ApprovedBudget.length == 0 ||
      this.state.EstimatedProjectCompletion.length == 0 ||
      this.state.CostCentre.length == 0 ||
      this.state.IsEstatesImplications == null  ||
      this.state.EstatesImplications.length == 0 ||
      this.state.ImplicationIT == null ||
      this.state.Implications.length == 0 ||
      
      this.state.ProjectManager.length == 0
      // this.state.ReviewerComment.length == 0 ||
      
    ) {
      alert("Please Complete the Details..!!");
    } else {
      const addCaptial : any = await sp.web.lists.getByTitle("Capital Project Authorization").items.add({
        IsUrgent: this.state.IsUrgent ,
        Purpose: this.state.Purpose,
        ProjectName: this.state.ProjectName,
        ProjectManagerId: this.state.ProjectManagerID,
        ProjectDescription: this.state.ProjectDescription,
        Location: this.state.Location,
        SpecificLocation: this.state.SpecificLocation,
        CapitalPlanCategory: this.state.CapitalPlanCategory,
        FundingSource: this.state.FundingSource,
        Approvalamount: this.state.Approvalamount,
        ProjectSponsor: this.state.ProjectSponsor,
        ContactEmail: this.state.ContactEmail,
        ApprovedBudget: this.state.ApprovedBudget,
        EstimatedProjectCompletion: this.state.EstimatedProjectCompletion,
        CostCentre: this.state.CostCentre,
        ImplicationIT: this.state.ImplicationIT,
        Implications: this.state.Implications,
        IsEstatesImplications: this.state.IsEstatesImplications ,
        EstatesImplications: this.state.EstatesImplications,
        Status: "P1 Pending",
        // ProjectManager: this.state.ProjectManager
        // ReviewComment: this.state.ReviewerComment,
        // ApproveCommet: this.state.Approvalamount,
      })
      .catch((error) => {
        console.log(error);
      });

      if(this.state.AllDocuments.length > 0){
        const libraryName = 'Capital Form Documents'; // Replace with your document library name
    
        for (let i = 0; i < this.state.AllDocuments.length; i++) {
          const file = this.state.AllDocuments[i];
    
          try {
              await sp.web.lists.getByTitle('Capital Form Documents').rootFolder.files.add(file.key.name, file.key, true).then(f =>
              f.file.getItem().then(Item => {
                Item.update({
                  DocumentType: file.text,
                  RequestIdId : addCaptial.data.ID
                }
                );
              })
            );
            console.log(`Uploaded: ${file.key.name}`);
          } catch (error) {
            console.error(`Error uploading ${file.key.name}:`, error);
          }
        }
      }

      if(this.state.DeleteDocuments.length>0){

        for(let i=0;i<this.state.DeleteDocuments.length;i++)
          {
          let id = this.state.DeleteDocuments[i];
          let web = Web(this.props.webURL);
    
          await web.lists.getByTitle("Capital Form Documents").items.getById(id.ID).delete()
          .then(i => {
            console.log(i);
          });
    
          }
      }
  
      this.setState({ AllDocuments: [] });// Clear the file input
      // this.GetDocumemtFromdocumentLibrary();
      this.GetCapitalForm();
      this.setState({ AddCapitalFormDialog: true });
      this.setState({ CaptialFormData: addCaptial });
    }

    this.setState({
      CaptialFormData : "",
      AddCapitalFormDialog: true,
      Title: "",
      IsUrgent: true,
      PurposeID: "",
      Purpose:"",
      ProjectName: "",
      ProjectDescription: "",
      Location: "",
      SpecificLocation: "",
      CapitalPlanCategory: "",
      FundingSource: "",
      Approvalamount: "",
      ProjectSponsorID: "",
      ProjectSponsorTitle: "",
      ContactEmail: "",
      ApprovedBudget: "",
      EstimatedProjectCompletion: "",
      CostCentre: "",
      IsEstatesImplications: true,
      EstatesImplications: "",
      ImplicationIT: true,
      Implications:"",
      Status: "",
      ReadReviewerComment: "",
      ApproverComment: "",
      ProjectManagerID: [],
      CapitalPurposelist: "",
      CapitalLocationlist: "",
      CapitalPlanCategorylist: "",
      CapitalFundingSourcelist: "",
      CapitalCostCenter: "",
      CapitalStatuslist: "",
      ProjectSponsor: "",
      ProjectManager: "",
      searchText: "",
      CapitalExportData: "",
      startDate : "",
      endDate:  "",
      EditIsUrgent :"",
      EditPurpose: "",
      EditProjectName: "",
      EditProjectDescription: "",
      EditLocation: "",
      EditSpecificLocation: "",
      EditCapitalPlanCategory: "",
      EditFundingSource: "",
      EditApprovalamount: "",
      EditProjectSponsor: "",
      EditProjectManagerID: "",
      EditProjectManager: "",
      EditContactEmail: "",
      EditApprovedBudget: "",
      EditEstimatedProjectCompletion: "",
      EditCostCentre: "",
      EditIsEstatesImplications: "",
      EditEstatesImplications: "",
      EditStatus: "",
      EditImplicationIT: "",
      EditImplications: "",
      EditFilterDialog : true,
      CurrentCapitalFormID :"",
      UpdateCapitalFormFilterDialog: "",
      EditReviewerComment: "",
      DeleteCurrentitem: "",
      DeleteFilterDialog :true,
      ReadFilterDialog: true,
      SearchData: "",
      ReadAllData: "",
      ReadIsUrgent: "",
      ReadPurpose:  "",
      ReadProjectName: "",
      ReadProjectDescription: "",
      ReadLocation: "",
      ReadSpecificLocation: "",
      ReadCapitalPlanCategory: "",
      ReadFundingSource: "",
      ReadProjectSponsor: "",
      ReadProjectManager: "",
      ReadContactEmail: "",
      ReadApprovedBudget: "", 
      ReadEstimatedProjectCompletion: "",
      ReadCostCentre: "",
      ReadIsEstatesImplications: "",
      ReadEstatesImplications: "",
      ReadStatus:"",
      ReadImplicationIT: "",
      ReadImplications: "",
      ReadApprovalamount: "",
      ReadProjectManagerID: "",
      ApprovalUser : false,
      ReviewerUser : false,
      ReturnFilter: true,
      ApproveFilter: true,
      IsReviewer: true,
      IsApproval: false,
      ReviewerComment: "",
      ReadReviewerID: "",
      ReviewerFilterDialog: false,
      ApprovalComment: "",
      ReadApprovalID : "",
      ReadApproverComment: "",
      CapitalFilterExportData : "",
      AllStatus: "",
      Pending : "",
      Approved : "",
      Rejected : "",
      ExportStatus : "",
      AllDocuments: [],
      GetAllDocuments: [],
      TempId:0,
      IncremenetState:1,
      DeleteDocuments:[],
      EditRequestId : "",
      CurrentDocumentID : "",
      ReadRequestId : "",
      DeleteDocument : ""
    });
} 

/* ChoiceForm CapitalFormAdd Call */
    public async GetCaptialFormChoiceItems() {
/*  ChoiceFiled1  */
          const choiceFieldName1 = "Purpose";
          const filed1 = await sp.web.lists.getByTitle("Capital Project Authorization").fields.getByInternalNameOrTitle(choiceFieldName1)();
          let ProjectPurposelist = [];
          filed1["Choices"].forEach(function (dname, i) {
            ProjectPurposelist.push({ key: dname, text: dname });
          });
          console.log(filed1);
          this.setState({ CapitalPurposelist : ProjectPurposelist });

/*  ChoiceFiled 2   */
          const choiceFieldName2 = "Location";
          const filed2 = await sp.web.lists.getByTitle("Capital Project Authorization").fields.getByInternalNameOrTitle(choiceFieldName2)();
          let ProjectLocationlist = [];
          filed2["Choices"].forEach(function (dname, i) {
            ProjectLocationlist.push({ key: dname, text: dname});
          });
          console.log(filed2);
          this.setState({ CapitalLocationlist: ProjectLocationlist });

/*  ChoiceFiled 3   */
          const choiceFieldName3 = "CapitalPlanCategory";
          const filed3 = await sp.web.lists.getByTitle("Capital Project Authorization").fields.getByInternalNameOrTitle(choiceFieldName3)();
          let ProjectCapitalPlanlist = [];
          filed3["Choices"].forEach(function (dname, i) {
            ProjectCapitalPlanlist.push({ key: dname , text: dname});
          });
          console.log(filed3);
          this.setState({ CapitalPlanCategorylist: ProjectCapitalPlanlist });

/*  ChoiceFiled 4  */
          const choiceFieldName4 = "FundingSource";
          const filed4 = await sp.web.lists.getByTitle("Capital Project Authorization").fields.getByInternalNameOrTitle(choiceFieldName4)();
          let ProjectFundingSource = [];
          filed4["Choices"].forEach(function (dname, i) {
            ProjectFundingSource.push({ key: dname, text: dname});
          });
          console.log(filed4);
          this.setState({ CapitalFundingSourcelist: ProjectFundingSource });

/*  ChoiceFiled 5  */
          const choiceFieldName5 = "CostCentre";
          const filed5 = await sp.web.lists.getByTitle("Capital Project Authorization").fields.getByInternalNameOrTitle(choiceFieldName5)();
          let ProjectCostCentre = [];
          filed5["Choices"].forEach(function (dname, i) {
            ProjectCostCentre.push({ key: dname, text: dname });
          });
          console.log(filed5);
          this.setState({ CapitalCostCenter : ProjectCostCentre });
          
 /*  ChoiceFiled 6  */
          const choiceFieldName6 = "Status";
          const filed6 = await sp.web.lists.getByTitle("Capital Project Authorization").fields.getByInternalNameOrTitle(choiceFieldName6)();
          let ProjectStatuslist = [];
          filed6["Choices"].forEach(function (dname, i) {
            ProjectStatuslist.push({ key: dname , text: dname });
          });
          console.log(filed6);
          this.setState({ CapitalStatuslist : ProjectStatuslist });
}

/* PeoplePicker */
  public _getPeoplePickerItems = async(items: any[]) => {

      if(items.length > 0 ) {
        this.setState({ ProjectManager: items[0].text });
        this.setState({ ProjectManagerID: items[0].id });
      }
      else {
        this.setState({ ProjectManager: "" });
        this.setState({ ProjectManagerID : "" });
      }
}

// /* Search Filter */  
//   public SearchCapital(searchText: string) : void {
//     const { CaptialFormData } = this.state;
    
//     const filteredItems = CaptialFormData.filter(
//       (item) => 
//         item.ProjectName.toLowerCase().includes(searchText.toLowerCase()) ||
//         item.ProjectManager.toLowerCase().includes(searchText.toLowerCase()) 
//     );

//     this.setState({ CaptialFormData : filteredItems });
// }

private  runexportfunction = (Test: string) => {
  const searchName = Test || '';
  this.setState({ searchText : searchName })
  this.Export_onFilter();
}


/* Search , StartDate, EndDate Filter */
  private async Export_onFilter() {
    if(this.state.searchText != 0 || this.state.startDate != 0 || this.state.endDate != 0 ) {
      let MySubTags = this.state.searchText.toLowerCase();

      let filterdData = this.state.CaptialFormData.filter((item) => {
        let Title = item.ProjectName.toLowerCase();
        let StartDate = moment(item.Created).format('DD MM YYYY');
        let EndDate = moment(item.Created).format('DD MM YYYY');
        console.log(this.state.startDate);
        console.log(this.state.endDate);

        let Record = (!MySubTags || Title.includes(MySubTags));

        let ExportStartdate = !this.state.startDate || (StartDate >= moment(this.state.startDate).format('DD MM YYYY'));

        let ExportEnddate = !this.state.endDate || (EndDate <= moment(this.state.endDate).format('DD MM YYYY'));

        console.log(this.state.startDate);
        console.log(this.state.endDate);
        console.log(StartDate);
        console.log(EndDate);

        return(
          Record && ExportStartdate && ExportEnddate
        );
      });
      this.setState({ CaptialFormData : filterdData });
    }

    else
    {
      console.log(this.state.CaptialFormData);
    }

    // this.GetCapitalForm();
  }

/* Read CapitalForm Data */
   public async GetReadCapitalForm(ID) {
    let readAlldata = this.state.CaptialFormData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    this.setState({
      IsUrgent: readAlldata[0].IsUrgent,
      ReadPurpose: readAlldata[0].Purpose,
      ReadProjectName: readAlldata[0].ProjectName,
      ReadProjectDescription: readAlldata[0].ProjectDescription,
      ReadLocation: readAlldata[0].Location,
      ReadSpecificLocation: readAlldata[0].SpecificLocation,
      ReadCapitalPlanCategory: readAlldata[0].CapitalPlanCategory,
      ReadFundingSource: readAlldata[0].FundingSource,
      ReadApprovalamount: readAlldata[0].Approvalamount,
      ReadProjectSponsor: readAlldata[0].ProjectSponsor,
      // ReadProjectManagerID: ReadCapitalAuthorization[0].ProjectManagerID,
      ReadProjectManager: readAlldata[0].ProjectManager,
      ReadContactEmail: readAlldata[0].ContactEmail,
      ReadApprovedBudget: readAlldata[0].ApprovedBudget,
      ReadEstimatedProjectCompletion: readAlldata[0].EstimatedProjectCompletion,
      ReadCostCentre: readAlldata[0].CostCentre,
      ReadIsEstatesImplications: readAlldata[0].IsEstatesImplications ,
      ReadEstatesImplications: readAlldata[0].EstatesImplications,
      ReadStatus: readAlldata[0].Status,
      ReadImplicationIT: readAlldata[0].ImplicationIT,
      ReadImplications: readAlldata[0].Implications,
      ReadReviewerID: readAlldata[0].ID,
      ReadApprovalID : readAlldata[0].ID,
      ReadReviewerComment: readAlldata[0].ReviewerComment,
      ReadApproverComment : readAlldata[0].ApproverComment,
      ReadRequestId: readAlldata[0].RequestId
    });
    
    this.setState({ ReadAllData : readAlldata, CurrentCapitalFormID : ID });
    console.log(readAlldata);
    this.GetCapitalForm();
}

/* Edit CapitalForm Data */
  public async GetEditCapitalForm(ID) {
    let EditCapitalAuthorization = this.state.CaptialFormData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditCapitalAuthorization);
    this.setState({
      IsUrgent: EditCapitalAuthorization[0].IsUrgent ,
      EditPurpose: EditCapitalAuthorization[0].Purpose,
      EditProjectName: EditCapitalAuthorization[0].ProjectName,
      EditProjectDescription: EditCapitalAuthorization[0].ProjectDescription,
      EditLocation: EditCapitalAuthorization[0].Location,
      EditSpecificLocation: EditCapitalAuthorization[0].SpecificLocation,
      EditCapitalPlanCategory: EditCapitalAuthorization[0].CapitalPlanCategory,
      EditFundingSource: EditCapitalAuthorization[0].FundingSource,
      EditApprovalamount: EditCapitalAuthorization[0].Approvalamount,
      EditProjectSponsor: EditCapitalAuthorization[0].ProjectSponsor,
      // EditProjectManagerID: EditCapitalAuthorization[0].ProjectManagerID,
      EditProjectManager: EditCapitalAuthorization[0].ProjectManager,
      EditContactEmail: EditCapitalAuthorization[0].ContactEmail,
      EditApprovedBudget: EditCapitalAuthorization[0].ApprovedBudget,
      EditEstimatedProjectCompletion: EditCapitalAuthorization[0].EstimatedProjectCompletion,
      EditCostCentre: EditCapitalAuthorization[0].CostCentre,
      EditIsEstatesImplications: EditCapitalAuthorization[0].IsEstatesImplications ,
      EditEstatesImplications: EditCapitalAuthorization[0].EstatesImplications,
      // EditStatus: EditCapitalAuthorization[0].Status,
      EditImplicationIT: EditCapitalAuthorization[0].ImplicationIT,
      EditImplications: EditCapitalAuthorization[0].Implications,
      EditRequestId: EditCapitalAuthorization[0].RequestId
    });
    console.log(
      this.state.IsUrgent,
      this.state.EditPurpose,
      this.state.EditProjectName,
      this.state.EditProjectDescription,
      this.state.EditLocation,
      this.state.EditSpecificLocation,
      this.state.EditCapitalPlanCategory,
      this.state.EditFundingSource,
      this.state.Approvalamount,
      this.state.EditProjectSponsor,
      this.state.EditProjectManager,
      this.state.EditContactEmail,
      this.state.EditApprovedBudget,
      this.state.EditIsEstatesImplications,
      this.state.EditEstatesImplications,
      // this.state.EditStatus,
      this.state.EditImplicationIT,
      this.state.EditImplications,
      this.state.EditRequestId
    );
    // this.GetDocumemtFromdocumentLibrary();
}

/* Update CapitalForm Data */
    public async UpdateCapitalForm(CurrentCapitalFormID) {
        const updateform  = await sp.web.lists.getByTitle("Capital Project Authorization").items.getById(CurrentCapitalFormID).update({
            IsUrgent: this.state.IsUrgent,
            Purpose: this.state.EditPurpose,
            ProjectName: this.state.EditProjectName,
            ProjectManagerId: this.state.ProjectManagerID ,
            ProjectDescription: this.state.EditProjectDescription,
            Location: this.state.EditLocation,
            SpecificLocation: this.state.EditSpecificLocation,
            CapitalPlanCategory: this.state.EditCapitalPlanCategory,
            FundingSource: this.state.EditFundingSource,
            Approvalamount: this.state.EditApprovalamount,
            ProjectSponsor: this.state.EditProjectSponsor,
            ContactEmail: this.state.EditContactEmail,
            ApprovedBudget: this.state.EditApprovedBudget,
            EstimatedProjectCompletion: this.state.EditEstimatedProjectCompletion,
            CostCentre: this.state.EditCostCentre,
            ImplicationIT: this.state.EditImplicationIT,
            Implications: this.state.EditImplications,
            IsEstatesImplications: this.state.EditIsEstatesImplications,
            EstatesImplications: this.state.EditEstatesImplications,
            // Status: this.state.EditStatus,
        }).catch((error) => {
          console.log(error);
        });

      
        
        if(this.state.AllDocuments.length > 0){
          const libraryName = 'Capital Form Documents'; // Replace with your document library name

          // const update = await sp.web.lists.getByTitle("Capital Form Documents").items.getById(CurrentDocumentID).update({
          //   RequestId : this.state.EditRequestId
          // }).catch((error) => {
          //   console.log(error);
          // });
      
          for (let i = 0; i < this.state.AllDocuments.length; i++) {
            const file = this.state.AllDocuments[i];
      
            try {
                await sp.web.lists.getByTitle('Capital Form Documents').rootFolder.files.add(file.key.name, file.key, true).then(f =>
                f.file.getItem().then(Item => {
                  Item.update({
                    DocumentType: file.text,
                    RequestIdId : CurrentCapitalFormID
                  }
                  );
                })
              );
              console.log(`Uploaded: ${file.key.name}`);
            } catch (error) {
              console.error(`Error uploading ${file.key.name}:`, error);
            }
          }
        }
  
        if(this.state.DeleteDocuments.length>0){
  
          for(let i=0;i<this.state.DeleteDocuments.length;i++)
            {
            let id = this.state.DeleteDocuments[i];
            let web = Web(this.props.webURL);
      
            await web.lists.getByTitle("Capital Form Documents").items.getById(id.ID).delete()
            .then(i => {
              console.log(i);
            });
      
            }
        }

        this.setState({ AllDocuments: [] });// Clear the file input
        // this.GetDocumemtFromdocumentLibrary();
        this.setState({ EditFilterDialog : true });
        this.GetCapitalForm();
    }

/* Delete CapitalForm Data */
    public async DeleteCapitalForm() {

      sp.web.lists.getByTitle('Capital Form Documents').items.select("*", 'Title', 'File/Name',"RequestId/Id").expand('File/Name',"RequestId").filter(`RequestId/Id eq ${this.state.DeleteCurrentitem}`).get()
      .then((data) => {
        let AllData = [];
          console.log(data);
        if (data.length > 0) {
          data.forEach((item, i) => {
            AllData.push({
              ID: item.Id ? item.Id : "",
              Filename: item.File.Name,
              DocumentType: item.DocumentType,
              Documentlink: item.ServerRedirectedEmbedUri,
              RequestId : item.RequestId ? item.RequestIdId : "",
              TempId: i
            });
              this.setState({TempId:i});
          });
          
          this.setState({ DeleteDocument: AllData });
          console.log(this.state.DeleteDocument);

          
          for(let i = 0;this.state.DeleteDocument.length>i;i++) {
            let deletedocument = this.state.DeleteDocument[i];
            sp.web.lists.getByTitle("Capital Form Documents").items.getById(deletedocument.ID).delete();
          }
          
        }
      })
      .catch((err) => {
        console.log(err);
      });


      const deletecapitalform = await sp.web.lists.getByTitle("Capital Project Authorization").items.getById(this.state.DeleteCurrentitem).delete();
      this.setState({ DeleteFilterDialog : true });
      this.GetCapitalForm();
  }

/* Start Date */
    public  handleStartDate(starDate : Date | null) : void {
    let startDate;  
    const matchesStartDate = startDate || startDate >= new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  
    this.GetCapitalForm();
    this.setState({ startDate : matchesStartDate });
    this.setState({ startDate: starDate });
  }

/* All Status View button */
  public async applyfilter(Status) {
    const { CapitalExportData } = this.state;
    let allStatusCount = 0;
    let pendingStatusCount = 0;
    let approvedStatusCount = 0;
    let rejectStatusCount = 0;

   
    CapitalExportData.forEach((item) => {
        allStatusCount++;
        
        if (item.Status === "P1 Pending" || item.Status === "P1 Approved") {
            pendingStatusCount++;
        }

        if (item.Status === "Approved" || item.Status === "P2 Approved") {
            approvedStatusCount++;
        }

        if (item.Status === "Rejected" || item.Status === "P1 Rejected" || item.Status === "P2 Rejected") {
            rejectStatusCount++;
        }
    });

    const FilterExportStatus = CapitalExportData.filter((item) => {
      if(Status == "allStatus") {
        const allStatus = item.Status;
        return allStatus;
      } else if(Status == "pendingStatus") {
        const pendingStatus = item.Status == "P1 Pending" || item.Status == "P1 Approved";
        return pendingStatus;
      } else if(Status == "approvedStatus") {
        const approvedStatus = item.Status == "Approved" || item.Status == "P2 Approved";
        return approvedStatus;
      } else if(Status == "rejectStatus") {
        const rejectStatus = item.Status == "Rejected" || item.Status == "P1 Rejected" || item.Status == "P2 Rejected";
        return rejectStatus;
      }
    });  
    this.setState({ CaptialFormData : FilterExportStatus });
    console.log(this.state.CaptialFormData);
  }


/* get All Documents From Document Library */
  public async GetDocumemtFromdocumentLibrary(ID) {

    sp.web.lists.getByTitle('Capital Form Documents').items.select("*", 'Title', 'File/Name',"RequestId/Id").expand('File/Name',"RequestId").filter(`RequestId/Id eq ${ID}`).get()
      .then((data) => {
        let AllData = [];
          console.log(data);
        if (data.length > 0) {
          data.forEach((item, i) => {
            AllData.push({
              ID: item.Id ? item.Id : "",
              Filename: item.File.Name,
              DocumentType: item.DocumentType,
              Documentlink: item.ServerRedirectedEmbedUri,
              RequestId : item.RequestId ? item.RequestIdId : "",
              TempId: i
            });
              this.setState({TempId:i});
          });
          this.setState({ GetAllDocuments: AllData  });
          console.log(this.state.GetAllDocuments);
        }
      })
      .catch((err) => {
        console.log(err);
      });
      this.GetCapitalForm();
  }

/* Get Attchments From Document Library */ 
  public GetAttchments(files, Doctype, RequestId) {
    // const field1: IFieldInfo = await sp.web.lists.getByTitle("MainEmployeeDetails").fields.getByInternalNameOrTitle("EmployeeDepartment")();
    let ProjectStatuslist = this.state.AllDocuments;
    let AllProjectStatuslist = this.state.GetAllDocuments;
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      ProjectStatuslist.push({ key: file, text: Doctype });
      AllProjectStatuslist.push({ID:"",
        Filename: file.name,
        DocumentType: Doctype,
        Documentlink: "",
        RequestId: RequestId ? RequestId : "",
        TempId:this.state.TempId+(i+1)
       });
       let test = this.state.TempId+(i+1);
       this.setState({TempId:test});
      }
    this.setState({ AllDocuments: ProjectStatuslist , GetAllDocuments:AllProjectStatuslist});
    console.log(this.state.AllDocuments);
    console.log(this.state.GetAllDocuments);
  }

/* Add new document and delete those document that user remove from webpage */
  public async handleUpload() {

    

   
  }

/* remove attchments from webpage and  */ 
  public removeAtttchments(tempid,Id,filename) {
    
    var array = this.state.GetAllDocuments;
    var array2 = this.state.AllDocuments;
 
   var index = array.findIndex(x => x.TempId === tempid);
   var index2 = array2.findIndex(x => x.key.name === filename);
    
    if (index !== -1) {
      array.splice(index, 1);
      this.setState({GetAllDocuments: array});
    }

    if (index2 !== -1) {
      array2.splice(index2, 1);
      this.setState({AllDocuments: array2});
    }

    if(Id){
      let deletedocuments = this.state.DeleteDocuments;
  
      deletedocuments.push(
        {
          ID:Id
        }
      );
      this.setState({DeleteDocuments:deletedocuments});
    }

    console.log(this.state.DeleteDocuments);
    console.log(this.state.AllDocuments);
    console.log(this.state.GetAllDocuments);
  }


// /* End Date */
// public handleEndDate(date : Date | null) : void {
//   const { endDate } = this.state;
//   const matchesEndDate = !endDate || endDate >= this.normalizeDate(endDate);
//   this.GetCapitalForm();
//   this.setState({ endDate : matchesEndDate });
//   this.setState({ endDate : date });
// }

}