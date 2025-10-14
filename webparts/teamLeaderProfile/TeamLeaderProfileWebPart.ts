import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";

import * as strings from 'TeamLeaderProfileWebPartStrings';

export interface ITeamLeaderProfileWebPartProps {
  description: string;
}

export default class TeamLeaderProfileWebPart extends BaseClientSideWebPart<ITeamLeaderProfileWebPartProps> {

  private teamMembers: any[] = [];
  private currentEditingId: number | null = null;
  private branchOptions: any[] = [];
  private branch1InternalName: string = "Branch1ENG";
  private branch2InternalName: string = "Branch2ENG";
  private branch3InternalName: string = "Branch3ENG";

  public async render(): Promise<void> {
    try {
      console.log("🔄 Starting Team Leader Profile render...");
      
      const sp = spfi().using(SPFx(this.context));
      const currentUser = await sp.web.currentUser();
      console.log("✅ Current team leader:", currentUser);

      // Get the internal names of the Branch ENG fields
      await this.getBranchFieldInternalNames(sp);

      // Fetch branch options from "Branch Names" list - USING SAME CONCEPT AS MYPROFILE
      try {
        this.branchOptions = await sp.web.lists.getByTitle("Branch Names").items
          .select("ID", "Title", "BranchNameENG")();
        console.log("✅ Branch options:", this.branchOptions);
      } catch (error) {
        console.error("❌ Error fetching branch options:", error);
        this.branchOptions = [];
      }
      
      // Build the select fields dynamically using internal names
      const selectFields = [
        "ID", 
        "Title", 
        "UserName/Id", 
        "UserName/Title",
        // Locked fields - using CORRECT internal names from debug
        "Code",
        "Department/Title",
        "SubSpeciality",        // "Sub Speciality" - CONFIRMED WORKING
        "ArmyStatus",           // "Army Status" - CONFIRMED WORKING
        "StartDate1",           // "Start Date 1" - CONFIRMED WORKING
        "StartDate2",           // "Start Date 2" - CONFIRMED WORKING 
        "StartDate3",           // "Start Date 3" - CONFIRMED WORKING
        "Shift1",
        "Shift2", 
        "Shift3",
        "Revenue",
        // Editable fields - using CORRECT internal names
        "Degree",
        "Name_x002d_EN",        // "Name - EN" - CONFIRMED WORKING
        "Name_x002d_AR",        // "Name - AR" - CONFIRMED WORKING
        "Specialty",
        "Email",
        "Phone",
        "DateofBirth",          // "Date of Birth" - CONFIRMED WORKING
        "MaritalStatus",        // "Marital Status" - CONFIRMED WORKING
        "Exclusive",
        "Bio",
        "BioEN",
        "CertificationInfo",    // "Certification Info" - CONFIRMED WORKING
        // Branch manager fields for filtering (Person fields)
        "Branch1Managedby/Id",
        "Branch1Managedby/Title",
        "Branch1Managedby/EMail",
        "Branch2Managedby/Id",
        "Branch2Managedby/Title", 
        "Branch2Managedby/EMail",
        "Branch3Managedby/Id",
        "Branch3Managedby/Title",
        "Branch3Managedby/EMail",
        // Branch ENG lookup fields using internal names
        `${this.branch1InternalName}/Id`,
        `${this.branch1InternalName}/Title`,
        `${this.branch2InternalName}/Id`,
        `${this.branch2InternalName}/Title`,
        `${this.branch3InternalName}/Id`,
        `${this.branch3InternalName}/Title`
      ];

      console.log("🔍 Select fields:", selectFields);
      
      // Get all employee data including branch manager fields (as Person fields)
      const items = await sp.web.lists.getByTitle("Employee Details").items
        .select(...selectFields)
        .expand("UserName", "Department", "Branch1Managedby", "Branch2Managedby", "Branch3Managedby", this.branch1InternalName, this.branch2InternalName, this.branch3InternalName)();
      
      console.log("✅ All items from list:", items);
      
      // Filter employees where current user matches any branch manager Person field
      const currentUserId = currentUser.Id;
      const currentUserEmail = currentUser.Email.toLowerCase();
      
      this.teamMembers = items.filter(item => {
        // Check each branch manager field (Person field)
        const branch1Manager = item.Branch1Managedby;
        const branch2Manager = item.Branch2Managedby;
        const branch3Manager = item.Branch3Managedby;
        
        // Check if current user is the manager by ID or Email
        const isManagedByBranch1 = branch1Manager && (
          branch1Manager.Id === currentUserId || 
          (branch1Manager.EMail && branch1Manager.EMail.toLowerCase() === currentUserEmail)
        );
        
        const isManagedByBranch2 = branch2Manager && (
          branch2Manager.Id === currentUserId || 
          (branch2Manager.EMail && branch2Manager.EMail.toLowerCase() === currentUserEmail)
        );
        
        const isManagedByBranch3 = branch3Manager && (
          branch3Manager.Id === currentUserId || 
          (branch3Manager.EMail && branch3Manager.EMail.toLowerCase() === currentUserEmail)
        );
        
        const isManaged = isManagedByBranch1 || isManagedByBranch2 || isManagedByBranch3;
        
        console.log(`🔍 Employee: ${item.Name_x002d_EN}, ` +
          `Branch1 Manager: ${branch1Manager ? `${branch1Manager.Title} (${branch1Manager.EMail})` : 'None'}, ` +
          `Branch2 Manager: ${branch2Manager ? `${branch2Manager.Title} (${branch2Manager.EMail})` : 'None'}, ` +
          `Branch3 Manager: ${branch3Manager ? `${branch3Manager.Title} (${branch3Manager.EMail})` : 'None'}, ` +
          `Current User: ${currentUser.Title} (${currentUser.Email}), ` +
          `Is Managed: ${isManaged}`
        );
        
        return isManaged;
      });
      
      console.log("✅ Filtered team members:", this.teamMembers);
      
      if (this.teamMembers.length === 0) {
        this.domElement.innerHTML = `
          <div style="font-family: Arial, sans-serif; padding: 20px;">
            <h2 style="color: #2c3e50;">Team Management</h2>
            <div style="padding: 15px; background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 4px; color: #856404;">
              <p><strong>No team members found.</strong></p>
              <p>You are not listed as a manager for any branch, or there are no employee records in the system.</p>
              <p><strong>Your User:</strong> ${currentUser.Title} (ID: ${currentUser.Id}, Email: ${currentUser.Email})</p>
              <p><strong>Total Employees in System:</strong> ${items.length}</p>
              <div style="margin-top: 10px; font-size: 14px;">
                <p><strong>Note:</strong> You need to be assigned as a Person in one of these fields for each employee:</p>
                <ul style="margin: 5px 0; padding-left: 20px;">
                  <li>Branch1Managedby</li>
                  <li>Branch2Managedby</li>
                  <li>Branch3Managedby</li>
                </ul>
              </div>
            </div>
          </div>
        `;
        return;
      }
      
      // Helper function to get branch display name - USING SAME CONCEPT AS MYPROFILE
      const getBranchDisplayName = (branch: any): string => {
        if (!branch) return '-';
        
        // If branch has a Title, use it
        if (branch.Title) return branch.Title;
        
        // If branch has ID but no Title, look it up in branchOptions
        if (branch.Id && this.branchOptions.length > 0) {
          const branchOption = this.branchOptions.find(b => b.ID === branch.Id);
          if (branchOption) {
            return branchOption.BranchNameENG || branchOption.Title || `Branch ${branch.Id}`;
          }
        }
        
        // Fallback
        return branch.Id ? `Branch ${branch.Id}` : '-';
      };

      // Helper function to safely display field values
      const displayValue = (value: any): string => {
        if (value === null || value === undefined || value === '') return '-';
        if (typeof value === 'object') {
          return value.Title || value.Value || JSON.stringify(value);
        }
        return value;
      };

      // Helper to format dates
      const formatDate = (dateValue: any): string => {
        if (!dateValue) return '-';
        try {
          const date = new Date(dateValue);
          return isNaN(date.getTime()) ? 'Invalid date' : date.toLocaleDateString();
        } catch {
          return String(dateValue);
        }
      };

      // Helper to format boolean values
      const formatBoolean = (value: any): string => {
        if (value === null || value === undefined) return '-';
        return value ? 'Yes' : 'No';
      };

      // Create HTML for the team management table
      let tableHTML = `
        <div style="font-family: Arial, sans-serif; max-width: 2200px; margin: 0 auto;">
          <h2 style="color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px;">Team Management</h2>
          <p style="color: #7f8c8d;">Welcome, <strong>${currentUser.Title}</strong>. You can manage schedule and branch information for your employees below.</p>
          <p style="color: #27ae60; font-weight: bold;">📊 Team Members: ${this.teamMembers.length}</p>
          <div style="padding: 10px; background: #e8f5e8; border-radius: 4px; margin-bottom: 15px; border: 1px solid #c8e6c9;">
            <strong>✅ Branch Filter Applied:</strong> Showing employees where you are assigned as manager in Branch1Managedby, Branch2Managedby, or Branch3Managedby Person fields.
          </div>
          
          <div style="overflow-x: auto; margin-top: 20px;">
            <table style="width: 100%; border-collapse: collapse; background: white; box-shadow: 0 1px 3px rgba(0,0,0,0.1); font-size: 10px;">
              <thead>
                <tr style="background: #34495e; color: white;">
                  <!-- Basic Information -->
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Code</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Name (EN)</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Name (AR)</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Degree</th>
                  
                  <!-- Contact Information -->
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Email</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Phone</th>
                  
                  <!-- Department & Role -->
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Department</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Specialty</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Sub Specialty</th>
                  
                  <!-- Personal Information -->
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Date of Birth</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Marital Status</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Army Status</th>
                  
                  <!-- Professional Information -->
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Exclusive</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Revenue</th>
                  
                  <!-- Branch ENG Information -->
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Branch 1 ENG</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Branch 2 ENG</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Branch 3 ENG</th>
                  
                  <!-- Bio Information -->
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Bio</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Bio EN</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Certification Info</th>
                  
                  <!-- Schedule Information -->
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Start Date 1</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Start Date 2</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Start Date 3</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Shift 1</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Shift 2</th>
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Shift 3</th>
                  
                  <!-- Actions -->
                  <th style="padding: 6px; text-align: left; border: 1px solid #2c3e50;">Actions</th>
                </tr>
              </thead>
              <tbody>
      `;
      
      // Add table rows for each team member
      this.teamMembers.forEach(member => {
        // Access branch fields using internal names
        const branch1ENG = member[this.branch1InternalName];
        const branch2ENG = member[this.branch2InternalName];
        const branch3ENG = member[this.branch3InternalName];
        
        tableHTML += `
          <tr style="border-bottom: 1px solid #ecf0f1;">
            <!-- Basic Information -->
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Code)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;"><strong>${displayValue(member.Name_x002d_EN)}</strong></td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Name_x002d_AR)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Degree)}</td>
            
            <!-- Contact Information -->
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Email)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Phone)}</td>
            
            <!-- Department & Role -->
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Department)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Specialty)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.SubSpeciality)}</td>
            
            <!-- Personal Information -->
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${formatDate(member.DateofBirth)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.MaritalStatus)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.ArmyStatus)}</td>
            
            <!-- Professional Information -->
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${formatBoolean(member.Exclusive)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Revenue)}</td>
            
            <!-- Branch ENG Information - USING SAME CONCEPT AS MYPROFILE -->
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${getBranchDisplayName(branch1ENG)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${getBranchDisplayName(branch2ENG)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${getBranchDisplayName(branch3ENG)}</td>
            
            <!-- Bio Information -->
            <td style="padding: 6px; border: 1px solid #ecf0f1; max-width: 100px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${displayValue(member.Bio)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1; max-width: 100px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${displayValue(member.BioEN)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1; max-width: 100px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${displayValue(member.CertificationInfo)}</td>
            
            <!-- Schedule Information -->
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${formatDate(member.StartDate1)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${formatDate(member.StartDate2)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${formatDate(member.StartDate3)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Shift1)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Shift2)}</td>
            <td style="padding: 6px; border: 1px solid #ecf0f1;">${displayValue(member.Shift3)}</td>
            
            <!-- Actions -->
            <td style="padding: 6px; border: 1px solid #ecf0f1;">
              <button class="editBtn" data-id="${member.ID}" style="padding: 4px 8px; background: #3498db; color: white; border: none; cursor: pointer; border-radius: 4px; font-size: 9px; width: 100%;">
                Edit Schedule & Branches
              </button>
            </td>
          </tr>
        `;
      });
      
      tableHTML += `
              </tbody>
            </table>
          </div>
          
          <!-- Edit Modal -->
          <div id="editModal" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000;">
            <div id="modalDialog" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 30px; border-radius: 8px; width: 95%; max-width: 900px; max-height: 90vh; overflow-y: auto;">
              <h3 style="color: #2c3e50; margin-bottom: 20px;">Edit Employee Schedule & Branch Information</h3>
              <div id="modalContent"></div>
              <div style="margin-top: 25px; text-align: right;">
                <button id="cancelBtn" style="padding: 10px 20px; background: #95a5a6; color: white; border: none; cursor: pointer; border-radius: 4px; margin-right: 10px;">Cancel</button>
                <button id="saveModalBtn" style="padding: 10px 20px; background: #27ae60; color: white; border: none; cursor: pointer; border-radius: 4px;">Save Changes</button>
              </div>
              <div id="modalMessage" style="margin-top: 15px; padding: 10px; border-radius: 4px;"></div>
            </div>
          </div>
        </div>
      `;
      
      this.domElement.innerHTML = tableHTML;
      
      // Add event handlers
      this.attachEventHandlers();
      
      console.log("✅ Team Leader Profile web part rendered successfully");
      
    } catch (error) {
      console.error("❌ Error in Team Leader Profile render:", error);
      this.domElement.innerHTML = `
        <div style="padding: 20px; color: #e74c3c; background: #fadbd8; border: 1px solid #e74c3c; border-radius: 4px;">
          <h2>Error</h2>
          <p><strong>Message:</strong> ${error.message}</p>
          <p>Check browser console for details.</p>
        </div>
      `;
    }
  }

  private async getBranchFieldInternalNames(sp: any): Promise<void> {
    try {
      // Get the internal name of the Branch1ENG field
      try {
        const branch1Field = await sp.web.lists.getByTitle("Employee Details").fields.getByTitle("Branch1ENG")();
        this.branch1InternalName = branch1Field.InternalName;
        console.log("✅ Branch1ENG field details:", {
          Title: branch1Field.Title,
          InternalName: branch1Field.InternalName,
          TypeAsString: branch1Field.TypeAsString,
          LookupList: branch1Field.LookupList,
          LookupField: branch1Field.LookupField
        });
      } catch (error) {
        console.error("❌ Error getting internal name for Branch1ENG, using default", error);
      }

      // Get the internal name of the Branch2ENG field
      try {
        const branch2Field = await sp.web.lists.getByTitle("Employee Details").fields.getByTitle("Branch2ENG")();
        this.branch2InternalName = branch2Field.InternalName;
        console.log("✅ Branch2ENG field details:", {
          Title: branch2Field.Title,
          InternalName: branch2Field.InternalName,
          TypeAsString: branch2Field.TypeAsString,
          LookupList: branch2Field.LookupList,
          LookupField: branch2Field.LookupField
        });
      } catch (error) {
        console.error("❌ Error getting internal name for Branch2ENG, using default", error);
      }

      // Get the internal name of the Branch3ENG field
      try {
        const branch3Field = await sp.web.lists.getByTitle("Employee Details").fields.getByTitle("Branch3ENG")();
        this.branch3InternalName = branch3Field.InternalName;
        console.log("✅ Branch3ENG field details:", {
          Title: branch3Field.Title,
          InternalName: branch3Field.InternalName,
          TypeAsString: branch3Field.TypeAsString,
          LookupList: branch3Field.LookupList,
          LookupField: branch3Field.LookupField
        });
      } catch (error) {
        console.error("❌ Error getting internal name for Branch3ENG, using default", error);
      }
    } catch (error) {
      console.error("❌ Error getting branch field internal names:", error);
    }
  }

  private attachEventHandlers(): void {
    const editButtons = this.domElement.querySelectorAll('.editBtn');
    const modal = document.getElementById('editModal') as HTMLDivElement;
    const modalDialog = document.getElementById('modalDialog') as HTMLDivElement;
    const modalContent = document.getElementById('modalContent') as HTMLDivElement;
    const modalMessage = document.getElementById('modalMessage') as HTMLDivElement;
    const cancelBtn = document.getElementById('cancelBtn') as HTMLButtonElement;
    const saveModalBtn = document.getElementById('saveModalBtn') as HTMLButtonElement;
    
    // Clear any existing event listeners first
    editButtons.forEach(button => {
      button.replaceWith(button.cloneNode(true));
    });
    
    const newEditButtons = this.domElement.querySelectorAll('.editBtn');
    
    newEditButtons.forEach(button => {
      button.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        
        const target = e.target as HTMLButtonElement;
        const itemId = parseInt(target.getAttribute('data-id')!);
        this.currentEditingId = itemId;
        
        const member = this.teamMembers.find(m => m.ID === itemId);
        if (!member) return;
        
        // Helper function for modal display
        const displayModalValue = (value: any): string => {
          if (value === null || value === undefined || value === '') return '';
          if (typeof value === 'object') return value.Title || '';
          return value;
        };

        // Generate branch dropdown options - USING SAME CONCEPT AS MYPROFILE
        const generateBranchOptions = (selectedBranch: any): string => {
          let options = '<option value="">-- Select Branch --</option>';
          this.branchOptions.forEach(branch => {
            const selected = selectedBranch && selectedBranch.Id === branch.ID ? 'selected' : '';
            const displayName = branch.BranchNameENG || branch.Title;
            options += `<option value="${branch.ID}" ${selected}>${displayName}</option>`;
          });
          return options;
        };

        // Access branch fields using internal names
        const branch1ENG = member[this.branch1InternalName];
        const branch2ENG = member[this.branch2InternalName];
        const branch3ENG = member[this.branch3InternalName];

        // Populate modal with edit form - showing all employee details
        modalContent.innerHTML = `
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 20px;">
            <!-- Column 1: Basic Information -->
            <div>
              <h4 style="color: #2c3e50; margin-bottom: 10px; border-bottom: 1px solid #bdc3c7; padding-bottom: 5px;">Basic Information</h4>
              <div style="margin-bottom: 8px;"><strong>Code:</strong> ${displayModalValue(member.Code)}</div>
              <div style="margin-bottom: 8px;"><strong>Name (EN):</strong> ${displayModalValue(member.Name_x002d_EN)}</div>
              <div style="margin-bottom: 8px;"><strong>Name (AR):</strong> ${displayModalValue(member.Name_x002d_AR)}</div>
              <div style="margin-bottom: 8px;"><strong>Degree:</strong> ${displayModalValue(member.Degree)}</div>
              <div style="margin-bottom: 8px;"><strong>Department:</strong> ${displayModalValue(member.Department)}</div>
            </div>
            
            <!-- Column 2: Contact & Professional -->
            <div>
              <h4 style="color: #2c3e50; margin-bottom: 10px; border-bottom: 1px solid #bdc3c7; padding-bottom: 5px;">Contact & Professional</h4>
              <div style="margin-bottom: 8px;"><strong>Email:</strong> ${displayModalValue(member.Email)}</div>
              <div style="margin-bottom: 8px;"><strong>Phone:</strong> ${displayModalValue(member.Phone)}</div>
              <div style="margin-bottom: 8px;"><strong>Specialty:</strong> ${displayModalValue(member.Specialty)}</div>
              <div style="margin-bottom: 8px;"><strong>Sub Specialty:</strong> ${displayModalValue(member.SubSpeciality)}</div>
              <div style="margin-bottom: 8px;"><strong>Revenue:</strong> ${displayModalValue(member.Revenue)}</div>
            </div>
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 20px;">
            <!-- Column 1: Personal Information -->
            <div>
              <h4 style="color: #2c3e50; margin-bottom: 10px; border-bottom: 1px solid #bdc3c7; padding-bottom: 5px;">Personal Information</h4>
              <div style="margin-bottom: 8px;"><strong>Date of Birth:</strong> ${member.DateofBirth ? new Date(member.DateofBirth).toLocaleDateString() : ''}</div>
              <div style="margin-bottom: 8px;"><strong>Marital Status:</strong> ${displayModalValue(member.MaritalStatus)}</div>
              <div style="margin-bottom: 8px;"><strong>Army Status:</strong> ${displayModalValue(member.ArmyStatus)}</div>
              <div style="margin-bottom: 8px;"><strong>Exclusive:</strong> ${member.Exclusive ? 'Yes' : 'No'}</div>
            </div>
            
            <!-- Column 2: Additional Information -->
            <div>
              <h4 style="color: #2c3e50; margin-bottom: 10px; border-bottom: 1px solid #bdc3c7; padding-bottom: 5px;">Additional Information</h4>
              <div style="margin-bottom: 8px;"><strong>Bio:</strong> ${displayModalValue(member.Bio) || 'N/A'}</div>
              <div style="margin-bottom: 8px;"><strong>Bio EN:</strong> ${displayModalValue(member.BioEN) || 'N/A'}</div>
              <div style="margin-bottom: 8px;"><strong>Certification Info:</strong> ${displayModalValue(member.CertificationInfo) || 'N/A'}</div>
            </div>
          </div>
          
          <div style="margin: 20px 0; padding: 15px; background: #e3f2fd; border-radius: 4px; border-left: 4px solid #2196f3;">
            <strong>🏢 Edit Branch Information:</strong>
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; margin-bottom: 20px;">
            <div>
              <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Branch 1 ENG:</label>
              <select id="modalBranch1ENG" style="padding: 8px; width: 100%; border: 1px solid #bdc3c7; border-radius: 4px;">
                ${generateBranchOptions(branch1ENG)}
              </select>
            </div>
            <div>
              <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Branch 2 ENG:</label>
              <select id="modalBranch2ENG" style="padding: 8px; width: 100%; border: 1px solid #bdc3c7; border-radius: 4px;">
                ${generateBranchOptions(branch2ENG)}
              </select>
            </div>
            <div>
              <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Branch 3 ENG:</label>
              <select id="modalBranch3ENG" style="padding: 8px; width: 100%; border: 1px solid #bdc3c7; border-radius: 4px;">
                ${generateBranchOptions(branch3ENG)}
              </select>
            </div>
          </div>
          
          <div style="margin: 20px 0; padding: 15px; background: #e3f2fd; border-radius: 4px; border-left: 4px solid #2196f3;">
            <strong>📅 Edit Schedule Information:</strong>
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">
            <div>
              <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Start Date 1:</label>
              <input type="date" id="modalStartDate1" value="${member.StartDate1 ? new Date(member.StartDate1).toISOString().split('T')[0] : ''}" style="padding: 8px; width: 100%; border: 1px solid #bdc3c7; border-radius: 4px;" />
            </div>
            <div>
              <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Start Date 2:</label>
              <input type="date" id="modalStartDate2" value="${member.StartDate2 ? new Date(member.StartDate2).toISOString().split('T')[0] : ''}" style="padding: 8px; width: 100%; border: 1px solid #bdc3c7; border-radius: 4px;" />
            </div>
            <div>
              <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Start Date 3:</label>
              <input type="date" id="modalStartDate3" value="${member.StartDate3 ? new Date(member.StartDate3).toISOString().split('T')[0] : ''}" style="padding: 8px; width: 100%; border: 1px solid #bdc3c7; border-radius: 4px;" />
            </div>
            <div>
              <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Shift 1:</label>
              <input type="text" id="modalShift1" value="${displayModalValue(member.Shift1)}" style="padding: 8px; width: 100%; border: 1px solid #bdc3c7; border-radius: 4px;" />
            </div>
            <div>
              <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Shift 2:</label>
              <input type="text" id="modalShift2" value="${displayModalValue(member.Shift2)}" style="padding: 8px; width: 100%; border: 1px solid #bdc3c7; border-radius: 4px;" />
            </div>
            <div>
              <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Shift 3:</label>
              <input type="text" id="modalShift3" value="${displayModalValue(member.Shift3)}" style="padding: 8px; width: 100%; border: 1px solid #bdc3c7; border-radius: 4px;" />
            </div>
          </div>
          
          <div style="margin-top: 20px; padding: 15px; background: #f8f9fa; border-radius: 4px; border: 1px solid #dee2e6;">
            <strong>⚠️ Note:</strong> You can edit schedule information (Start Dates and Shifts) and branch assignments (Branch ENG fields). Other employee details are managed by HR.
          </div>
        `;
        
        modal.style.display = 'block';
        modalMessage.innerHTML = '';
      });
    });
    
    // Cancel button handler
    cancelBtn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      modal.style.display = 'none';
      this.currentEditingId = null;
    });
    
    // Save button handler - UPDATED WITH BRANCH FUNCTIONALITY
    saveModalBtn.addEventListener('click', async (e) => {
      e.preventDefault();
      e.stopPropagation();
      
      if (!this.currentEditingId) return;
      
      try {
        const sp = spfi().using(SPFx(this.context));
        
        saveModalBtn.disabled = true;
        saveModalBtn.textContent = 'Saving...';
        saveModalBtn.style.background = '#95a5a6';
        modalMessage.innerHTML = 'Saving changes...';
        modalMessage.style.color = '#2980b9';
        modalMessage.style.backgroundColor = '#d6eaf8';
        modalMessage.style.border = '1px solid #3498db';
        
        // Prepare update data - schedule fields and branch ENG fields
        const updatedData: any = {};
        
        // Date fields
        const startDate1Value = (document.getElementById('modalStartDate1') as HTMLInputElement).value;
        const startDate2Value = (document.getElementById('modalStartDate2') as HTMLInputElement).value;
        const startDate3Value = (document.getElementById('modalStartDate3') as HTMLInputElement).value;
        
        updatedData.StartDate1 = startDate1Value ? new Date(startDate1Value).toISOString() : null;
        updatedData.StartDate2 = startDate2Value ? new Date(startDate2Value).toISOString() : null;
        updatedData.StartDate3 = startDate3Value ? new Date(startDate3Value).toISOString() : null;
        
        // Shift fields
        updatedData.Shift1 = (document.getElementById('modalShift1') as HTMLInputElement).value || null;
        updatedData.Shift2 = (document.getElementById('modalShift2') as HTMLInputElement).value || null;
        updatedData.Shift3 = (document.getElementById('modalShift3') as HTMLInputElement).value || null;
        
        // Branch ENG fields (lookup fields) - USING SAME CONCEPT AS MYPROFILE
        const branch1ENGValue = (document.getElementById('modalBranch1ENG') as HTMLSelectElement).value;
        const branch2ENGValue = (document.getElementById('modalBranch2ENG') as HTMLSelectElement).value;
        const branch3ENGValue = (document.getElementById('modalBranch3ENG') as HTMLSelectElement).value;

        console.log("🔍 Branch selection values:", {
          branch1ENGValue,
          branch2ENGValue,
          branch3ENGValue
        });

        console.log("🔍 Internal names being used:", {
          branch1InternalName: this.branch1InternalName,
          branch2InternalName: this.branch2InternalName,
          branch3InternalName: this.branch3InternalName
        });

        // Update branch lookup fields with proper format
        if (branch1ENGValue) {
          updatedData[`${this.branch1InternalName}Id`] = parseInt(branch1ENGValue);
        } else {
          updatedData[`${this.branch1InternalName}Id`] = null;
        }

        if (branch2ENGValue) {
          updatedData[`${this.branch2InternalName}Id`] = parseInt(branch2ENGValue);
        } else {
          updatedData[`${this.branch2InternalName}Id`] = null;
        }

        if (branch3ENGValue) {
          updatedData[`${this.branch3InternalName}Id`] = parseInt(branch3ENGValue);
        } else {
          updatedData[`${this.branch3InternalName}Id`] = null;
        }
        
        console.log("🔄 Final update data being sent:", updatedData);
        
        // Update the item
        const result = await sp.web.lists.getByTitle("Employee Details").items.getById(this.currentEditingId).update(updatedData);
        console.log("✅ Update result:", result);
        
        // Verify the update by fetching the item again
        const updatedItem = await sp.web.lists.getByTitle("Employee Details").items.getById(this.currentEditingId)
          .select(`${this.branch1InternalName}/Id`, `${this.branch1InternalName}/Title`, 
                  `${this.branch2InternalName}/Id`, `${this.branch2InternalName}/Title`,
                  `${this.branch3InternalName}/Id`, `${this.branch3InternalName}/Title`)
          .expand(this.branch1InternalName, this.branch2InternalName, this.branch3InternalName)();
        
        console.log("✅ Verified updated item branch fields:", {
          branch1ENG: updatedItem[this.branch1InternalName],
          branch2ENG: updatedItem[this.branch2InternalName],
          branch3ENG: updatedItem[this.branch3InternalName]
        });
        
        modalMessage.innerHTML = '✅ Schedule and branch information updated successfully!';
        modalMessage.style.color = '#27ae60';
        modalMessage.style.backgroundColor = '#d5f4e6';
        modalMessage.style.border = '1px solid #2ecc71';
        
        // Close modal and refresh after 2 seconds
        setTimeout(() => {
          modal.style.display = 'none';
          this.currentEditingId = null;
          this.render(); // Refresh the table
        }, 2000);
        
      } catch (error) {
        console.error("❌ Save error:", error);
        
        // More detailed error information
        let errorMessage = `❌ Error saving: ${error.message}`;
        if (error.data) {
          errorMessage += `<br>Details: ${JSON.stringify(error.data)}`;
        }
        if (error.status) {
          errorMessage += `<br>Status: ${error.status}`;
        }
        
        modalMessage.innerHTML = errorMessage;
        modalMessage.style.color = '#e74c3c';
        modalMessage.style.backgroundColor = '#fadbd8';
        modalMessage.style.border = '1px solid #e74c3c';
        saveModalBtn.textContent = 'Save Changes';
        saveModalBtn.disabled = false;
        saveModalBtn.style.background = '#27ae60';
      }
    });
    
    // Close modal when clicking outside
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        modal.style.display = 'none';
        this.currentEditingId = null;
      }
    });

    // Prevent modal from closing when clicking inside the dialog
    modalDialog.addEventListener('click', (e) => {
      e.stopPropagation();
    });
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    // Empty
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
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