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
import "@pnp/sp/attachments";

import * as strings from "MyProfileWebPartStrings";

export interface IMyProfileWebPartProps {
  description: string;
}

export default class MyProfileWebPart extends BaseClientSideWebPart<IMyProfileWebPartProps> {

  private branch1InternalName: string = "Branch1ENG";
  private branch2InternalName: string = "Branch2ENG";
  private branch3InternalName: string = "Branch3ENG";
  private branchOptions: any[] = [];

  public async render(): Promise<void> {
    try {
      console.log("🔄 Starting render with corrected field names...");
      
      const sp = spfi().using(SPFx(this.context));
      const currentUser = await sp.web.currentUser();
      console.log("✅ Current user:", currentUser);

      // Get the internal names of the Branch ENG fields
      await this.getBranchFieldInternalNames(sp);

      // Fetch branch options from "Branch Names" list - LIKE TEAM LEADER PROFILE
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
        "Picture",
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
        // Branch ENG lookup fields using internal names
        `${this.branch1InternalName}/Id`,
        `${this.branch1InternalName}/Title`,
        `${this.branch2InternalName}/Id`,
        `${this.branch2InternalName}/Title`,
        `${this.branch3InternalName}/Id`,
        `${this.branch3InternalName}/Title`
      ];

      console.log("🔍 Select fields:", selectFields);
      
      // Get employee data
      const items = await sp.web.lists.getByTitle("Employee Details").items
        .select(...selectFields)
        .expand("UserName", "Department", this.branch1InternalName, this.branch2InternalName, this.branch3InternalName)();
      
      console.log("✅ All items from list:", items);
      
      // Filter by UserName person field
      const userItems = items.filter(item => 
        item.UserName && item.UserName.Id === currentUser.Id
      );
      
      console.log("✅ Filtered user items:", userItems);
      
      if (userItems.length === 0) {
        this.domElement.innerHTML = `
          <div>
            <h2>Employee Profile</h2>
            <p>No employee record found for ${currentUser.Title}.</p>
            <p>Please contact HR to create your employee record.</p>
          </div>
        `;
        return;
      }
      
      const emp = userItems[0];
      console.log("✅ Employee data to display:", emp);

      // Get attachments for Certification Info
      let attachments: any[] = [];
      try {
        attachments = await sp.web.lists.getByTitle("Employee Details").items.getById(emp.ID).attachmentFiles();
        console.log("✅ Attachments:", attachments);
      } catch (error) {
        console.error("❌ Error fetching attachments:", error);
      }

      // Access branch fields using internal names
      const branch1ENG = emp[this.branch1InternalName];
      const branch2ENG = emp[this.branch2InternalName];
      const branch3ENG = emp[this.branch3InternalName];

      console.log("🔍 Branch fields:", {
        branch1ENG,
        branch2ENG,
        branch3ENG,
        branch1InternalName: this.branch1InternalName,
        branch2InternalName: this.branch2InternalName,
        branch3InternalName: this.branch3InternalName
      });

      // Helper function to get branch display name - FIXED FOR NULL TITLES
      const getBranchDisplayName = (branch: any): string => {
        if (!branch) return 'Not set';
        
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
        return branch.Id ? `Branch ${branch.Id}` : 'Not set';
      };
      
      // Helper function to safely display field values
      const displayValue = (value: any): string => {
        if (value === null || value === undefined || value === '') return 'Not set';
        if (typeof value === 'object') {
          return value.Title || value.Value || JSON.stringify(value);
        }
        if (typeof value === 'boolean') {
          return value ? 'Yes' : 'No';
        }
        return value;
      };

      // Helper to format dates
      const formatDate = (dateValue: any): string => {
        if (!dateValue) return 'Not set';
        try {
          const date = new Date(dateValue);
          return isNaN(date.getTime()) ? 'Invalid date' : date.toLocaleDateString();
        } catch {
          return String(dateValue);
        }
      };
      
      // Create HTML for the form
      let formHTML = `
        <div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto;">
          <h2 style="color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px;">Employee Profile</h2>
          <p><strong>Employee:</strong> ${currentUser.Title}</p>
          <p><strong>Email:</strong> ${currentUser.Email}</p>
          <hr style="margin: 20px 0; border: 1px solid #eee;">
          
          <h3 style="color: #e74c3c; margin-bottom: 15px;">🔒 Locked Information (HR & Team Leader Managed)</h3>
      `;
      
      // REORDERED: Picture first
      formHTML += `
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Picture:</strong><br/>
          ${emp.Picture ? `<img src="${emp.Picture}" alt="Profile" style="max-width: 150px; max-height: 150px; border-radius: 4px; margin-top: 8px;" />` : '<span style="color: #495057;">Not set</span>'}
        </div>
      `;
      
      // Code
      formHTML += `
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Employee Code:</strong><br/>
          <span style="color: #495057;">${displayValue(emp.Code)}</span>
        </div>
      `;
      
      // Branch fields
      formHTML += `
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Branch 1:</strong><br/>
          <span style="color: #495057;">${getBranchDisplayName(branch1ENG)}</span>
        </div>
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Branch 2:</strong><br/>
          <span style="color: #495057;">${getBranchDisplayName(branch2ENG)}</span>
        </div>
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Branch 3:</strong><br/>
          <span style="color: #495057;">${getBranchDisplayName(branch3ENG)}</span>
        </div>
      `;
      
      // Start Dates
      formHTML += `
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Start Date 1:</strong><br/>
          <span style="color: #495057;">${formatDate(emp.StartDate1)}</span>
        </div>
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Start Date 2:</strong><br/>
          <span style="color: #495057;">${formatDate(emp.StartDate2)}</span>
        </div>
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Start Date 3:</strong><br/>
          <span style="color: #495057;">${formatDate(emp.StartDate3)}</span>
        </div>
      `;
      
      // Shifts
      formHTML += `
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Shift 1:</strong><br/>
          <span style="color: #495057;">${displayValue(emp.Shift1)}</span>
        </div>
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Shift 2:</strong><br/>
          <span style="color: #495057;">${displayValue(emp.Shift2)}</span>
        </div>
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Shift 3:</strong><br/>
          <span style="color: #495057;">${displayValue(emp.Shift3)}</span>
        </div>
      `;
      
      // REORDERED: Department, Sub Speciality, Army Status after shifts
      formHTML += `
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Department:</strong><br/>
          <span style="color: #495057;">${displayValue(emp.Department)}</span>
        </div>
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Sub Speciality:</strong><br/>
          <span style="color: #495057;">${displayValue(emp.SubSpeciality)}</span>
        </div>
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Army Status:</strong><br/>
          <span style="color: #495057;">${displayValue(emp.ArmyStatus)}</span>
        </div>
      `;
      
      // Revenue
      formHTML += `
        <div style="margin: 10px 0; padding: 12px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
          <strong>Revenue:</strong><br/>
          <span style="color: #495057;">${displayValue(emp.Revenue)}</span>
        </div>
      `;
      
      // Add editable fields section
      formHTML += `
        <hr style="margin: 30px 0; border: 1px solid #eee;">
        <h3 style="color: #27ae60; margin-bottom: 15px;">✏️ Editable Information (You Can Update)</h3>
      `;
      
      // Degree
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Degree:</label>
          <input type="text" id="degreeField" value="${displayValue(emp.Degree)}" style="padding: 10px; width: 100%; max-width: 400px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px;" />
        </div>
      `;
      
      // Name - EN
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Name - EN:</label>
          <input type="text" id="nameENField" value="${displayValue(emp.Name_x002d_EN)}" style="padding: 10px; width: 100%; max-width: 400px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px;" />
        </div>
      `;
      
      // Name - AR
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Name - AR:</label>
          <input type="text" id="nameARField" value="${displayValue(emp.Name_x002d_AR)}" style="padding: 10px; width: 100%; max-width: 400px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px;" />
        </div>
      `;
      
      // Specialty
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Specialty:</label>
          <input type="text" id="specialtyField" value="${displayValue(emp.Specialty)}" style="padding: 10px; width: 100%; max-width: 400px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px;" />
        </div>
      `;
      
      // Email
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Email:</label>
          <input type="email" id="emailField" value="${displayValue(emp.Email)}" style="padding: 10px; width: 100%; max-width: 400px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px;" />
        </div>
      `;
      
      // Phone
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Phone:</label>
          <input type="tel" id="phoneField" value="${displayValue(emp.Phone)}" style="padding: 10px; width: 100%; max-width: 400px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px;" />
        </div>
      `;
      
      // Date of Birth
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Date of Birth:</label>
          <input type="date" id="dobField" value="${emp.DateofBirth ? new Date(emp.DateofBirth).toISOString().split('T')[0] : ''}" style="padding: 10px; width: 100%; max-width: 400px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px;" />
        </div>
      `;
      
      // Marital Status
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Marital Status:</label>
          <select id="maritalStatusField" style="padding: 10px; width: 100%; max-width: 400px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px; background: white;">
            <option value="">Select Status</option>
            <option value="Single" ${emp.MaritalStatus === 'Single' ? 'selected' : ''}>Single</option>
            <option value="Married" ${emp.MaritalStatus === 'Married' ? 'selected' : ''}>Married</option>
            <option value="Divorced" ${emp.MaritalStatus === 'Divorced' ? 'selected' : ''}>Divorced</option>
            <option value="Widowed" ${emp.MaritalStatus === 'Widowed' ? 'selected' : ''}>Widowed</option>
          </select>
        </div>
      `;
      
      // Exclusive
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Exclusive:</label>
          <select id="exclusiveField" style="padding: 10px; width: 100%; max-width: 400px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px; background: white;">
            <option value="false" ${emp.Exclusive === false ? 'selected' : ''}>No</option>
            <option value="true" ${emp.Exclusive === true ? 'selected' : ''}>Yes</option>
          </select>
        </div>
      `;
      
      // Bio
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Bio (AR):</label>
          <textarea id="bioField" style="padding: 10px; width: 100%; max-width: 500px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px; height: 100px; font-family: Arial, sans-serif;">${displayValue(emp.Bio)}</textarea>
        </div>
      `;
      
      // BioEN
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Bio (EN):</label>
          <textarea id="bioENField" style="padding: 10px; width: 100%; max-width: 500px; border: 1px solid #bdc3c7; border-radius: 4px; font-size: 14px; height: 100px; font-family: Arial, sans-serif;">${displayValue(emp.BioEN)}</textarea>
        </div>
      `;
      
      // Certification Info as ATTACHMENTS
      formHTML += `
        <div style="margin: 15px 0;">
          <label style="display: block; margin-bottom: 5px; font-weight: bold; color: #2c3e50;">Certification Info (Attachments):</label>
          <div id="attachmentsList" style="margin-bottom: 15px;">
      `;
      
      if (attachments.length > 0) {
        attachments.forEach(attachment => {
          formHTML += `
            <div style="display: flex; align-items: center; margin-bottom: 8px; padding: 8px; background: #f8f9fa; border-radius: 4px;">
              <a href="${attachment.ServerRelativeUrl}" target="_blank" style="flex: 1; color: #3498db; text-decoration: none;">
                📎 ${attachment.FileName}
              </a>
              <button type="button" class="deleteAttachment" data-filename="${attachment.FileName}" style="background: #e74c3c; color: white; border: none; border-radius: 3px; padding: 4px 8px; cursor: pointer; margin-left: 10px;">
                Delete
              </button>
            </div>
          `;
        });
      } else {
        formHTML += `<div style="color: #7f8c8d; font-style: italic;">No attachments uploaded</div>`;
      }
      
      formHTML += `
          </div>
          <div>
            <input type="file" id="newAttachments" multiple style="margin-bottom: 10px;" />
            <button type="button" id="uploadAttachments" style="padding: 8px 16px; background: #3498db; color: white; border: none; border-radius: 4px; cursor: pointer;">Upload Files</button>
            <div style="font-size: 12px; color: #7f8c8d; margin-top: 5px;">You can upload multiple certification files</div>
          </div>
        </div>
      `;
      
      // Add save button and message area
      formHTML += `
        <br/>
        <button id="saveBtn" style="padding: 12px 30px; background: #3498db; color: white; border: none; cursor: pointer; border-radius: 4px; font-size: 16px; font-weight: bold; margin-top: 20px;">Save Changes</button>
        <div id="message" style="margin-top: 15px; padding: 12px; border-radius: 4px;"></div>
      </div>
      `;
      
      this.domElement.innerHTML = formHTML;
      
      // Add event handlers
      this.attachEventHandlers(sp, emp.ID);
      
      console.log("✅ Web part rendered successfully");
      
    } catch (error) {
      console.error("❌ Error in render:", error);
      this.domElement.innerHTML = `
        <div style="padding: 20px; color: #e74c3c; background: #fadbd8; border: 1px solid #e74c3c; border-radius: 4px;">
          <h2>Error</h2>
          <p><strong>Message:</strong> ${error.message}</p>
          <p>Check browser console for details.</p>
        </div>
      `;
    }
  }

  private attachEventHandlers(sp: any, itemId: number): void {
    const saveBtn = document.getElementById('saveBtn') as HTMLButtonElement;
    const messageDiv = document.getElementById('message') as HTMLDivElement;
    const uploadBtn = document.getElementById('uploadAttachments') as HTMLButtonElement;
    
    // Save button handler
    saveBtn.addEventListener('click', async () => {
      try {
        saveBtn.disabled = true;
        saveBtn.textContent = 'Saving...';
        saveBtn.style.background = '#95a5a6';
        messageDiv.innerHTML = 'Saving changes...';
        messageDiv.style.color = '#2980b9';
        messageDiv.style.backgroundColor = '#d6eaf8';
        messageDiv.style.border = '1px solid #3498db';
        
        // Prepare update data - only include editable fields (BRANCH FIELDS ARE VIEW-ONLY, NOT INCLUDED)
        const updatedData: any = {};
        
        // Text fields
        updatedData.Degree = (document.getElementById('degreeField') as HTMLInputElement).value || null;
        updatedData.Name_x002d_EN = (document.getElementById('nameENField') as HTMLInputElement).value || null;
        updatedData.Name_x002d_AR = (document.getElementById('nameARField') as HTMLInputElement).value || null;
        updatedData.Specialty = (document.getElementById('specialtyField') as HTMLInputElement).value || null;
        updatedData.Email = (document.getElementById('emailField') as HTMLInputElement).value || null;
        updatedData.Phone = (document.getElementById('phoneField') as HTMLInputElement).value || null;
        
        // Date field
        const dobValue = (document.getElementById('dobField') as HTMLInputElement).value;
        updatedData.DateofBirth = dobValue ? new Date(dobValue).toISOString() : null;
        
        // Choice fields
        updatedData.MaritalStatus = (document.getElementById('maritalStatusField') as HTMLSelectElement).value || null;
        
        // Yes/No field
        updatedData.Exclusive = (document.getElementById('exclusiveField') as HTMLSelectElement).value === 'true';
        
        // Text areas
        updatedData.Bio = (document.getElementById('bioField') as HTMLTextAreaElement).value || null;
        updatedData.BioEN = (document.getElementById('bioENField') as HTMLTextAreaElement).value || null;
        
        console.log("🔄 Updating data:", updatedData);
        
        // Update the item
        await sp.web.lists.getByTitle("Employee Details").items.getById(itemId).update(updatedData);
        
        messageDiv.innerHTML = '✅ Changes saved successfully!';
        messageDiv.style.color = '#27ae60';
        messageDiv.style.backgroundColor = '#d5f4e6';
        messageDiv.style.border = '1px solid #2ecc71';
        saveBtn.textContent = 'Save Changes';
        saveBtn.disabled = false;
        saveBtn.style.background = '#3498db';
        
        // Refresh the data after 2 seconds
        setTimeout(() => {
          this.render();
        }, 2000);
        
      } catch (error) {
        console.error("❌ Save error:", error);
        messageDiv.innerHTML = `❌ Error saving: ${error.message}`;
        messageDiv.style.color = '#e74c3c';
        messageDiv.style.backgroundColor = '#fadbd8';
        messageDiv.style.border = '1px solid #e74c3c';
        saveBtn.textContent = 'Save Changes';
        saveBtn.disabled = false;
        saveBtn.style.background = '#3498db';
      }
    });
    
    // Upload attachments handler
    uploadBtn.addEventListener('click', async () => {
      const fileInput = document.getElementById('newAttachments') as HTMLInputElement;
      
      if (!fileInput.files || fileInput.files.length === 0) {
        this.showMessage(messageDiv, 'Please select at least one file to upload', 'warning');
        return;
      }
      
      try {
        uploadBtn.disabled = true;
        uploadBtn.textContent = 'Uploading...';
        this.showMessage(messageDiv, 'Uploading files...', 'info');
        
        // Upload each file
        for (let i = 0; i < fileInput.files.length; i++) {
          const file = fileInput.files[i];
          console.log(`📤 Uploading file: ${file.name}`);
          
          await sp.web.lists.getByTitle("Employee Details").items.getById(itemId).attachmentFiles.add(file.name, file);
        }
        
        this.showMessage(messageDiv, `✅ Successfully uploaded ${fileInput.files.length} file(s)`, 'success');
        uploadBtn.textContent = 'Upload Files';
        uploadBtn.disabled = false;
        fileInput.value = ''; // Clear file input
        
        // Refresh attachments list after 2 seconds
        setTimeout(() => {
          this.render();
        }, 2000);
        
      } catch (error) {
        console.error("❌ Upload error:", error);
        this.showMessage(messageDiv, `❌ Error uploading files: ${error.message}`, 'error');
        uploadBtn.textContent = 'Upload Files';
        uploadBtn.disabled = false;
      }
    });
    
    // Delete attachment handlers
    const deleteButtons = this.domElement.querySelectorAll('.deleteAttachment');
    deleteButtons.forEach(button => {
      button.addEventListener('click', async (e) => {
        const target = e.target as HTMLButtonElement;
        const fileName = target.getAttribute('data-filename');
        
        if (!fileName) return;
        
        if (confirm(`Are you sure you want to delete "${fileName}"?`)) {
          try {
            target.disabled = true;
            target.textContent = 'Deleting...';
            target.style.background = '#95a5a6';
            
            await sp.web.lists.getByTitle("Employee Details").items.getById(itemId).attachmentFiles.getByName(fileName).delete();
            
            this.showMessage(messageDiv, `✅ File "${fileName}" deleted successfully`, 'success');
            
            // Refresh attachments list after 1 second
            setTimeout(() => {
              this.render();
            }, 1000);
            
          } catch (error) {
            console.error("❌ Delete error:", error);
            this.showMessage(messageDiv, `❌ Error deleting file: ${error.message}`, 'error');
            target.textContent = 'Delete';
            target.disabled = false;
            target.style.background = '#e74c3c';
          }
        }
      });
    });
  }
  
  private showMessage(messageDiv: HTMLDivElement, message: string, type: 'info' | 'success' | 'warning' | 'error'): void {
    messageDiv.innerHTML = message;
    
    switch (type) {
      case 'info':
        messageDiv.style.color = '#2980b9';
        messageDiv.style.backgroundColor = '#d6eaf8';
        messageDiv.style.border = '1px solid #3498db';
        break;
      case 'success':
        messageDiv.style.color = '#27ae60';
        messageDiv.style.backgroundColor = '#d5f4e6';
        messageDiv.style.border = '1px solid #2ecc71';
        break;
      case 'warning':
        messageDiv.style.color = '#f39c12';
        messageDiv.style.backgroundColor = '#fef5e7';
        messageDiv.style.border = '1px solid #f1c40f';
        break;
      case 'error':
        messageDiv.style.color = '#e74c3c';
        messageDiv.style.backgroundColor = '#fadbd8';
        messageDiv.style.border = '1px solid #e74c3c';
        break;
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