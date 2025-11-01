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
      console.log("🔄 Starting View Profile render...");
      
      const sp = spfi().using(SPFx(this.context));
      const currentUser = await sp.web.currentUser();
      console.log("✅ Current user:", currentUser);

      // Get the internal names of the Branch ENG fields
      await this.getBranchFieldInternalNames(sp);

      // Fetch branch options from "Branch Names" list
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
        // All fields for view only
        "Code",
        "Department/Title",
        "SubSpeciality",
        "ArmyStatus",
        "StartDate1",
        "StartDate2", 
        "StartDate3",
        "Shift1",
        "Shift2", 
        "Shift3",
        "Picture",
        "Revenue",
        "Specialty",
        "Exclusive",
        "Degree",
        "Name_x002d_EN",
        "Name_x002d_AR",
        "Email",
        "Phone",
        "DateofBirth",
        "MaritalStatus",
        "Bio",
        "BioEN",
        // Title fields (AR and EN)
        "Title_x002d_AR",
        "TitleEN",
        // Branch ENG lookup fields
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
          <div class="profile-container">
            <div class="profile-header">
              <h1>My Profile</h1>
            </div>
            <div class="error-message">
              <h2>No Employee Record Found</h2>
              <p>No employee record found for ${currentUser.Title}.</p>
              <p>Please contact HR to create your employee record.</p>
            </div>
          </div>
        `;
        return;
      }
      
      const emp = userItems[0];
      console.log("✅ Employee data to display:", emp);
      console.log("🖼️ Picture field value:", emp.Picture);

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

      // Helper function to get branch display name
      const getBranchDisplayName = (branch: any): string => {
        if (!branch) return 'Not set';
        if (branch.Title) return branch.Title;
        if (branch.Id && this.branchOptions.length > 0) {
          const branchOption = this.branchOptions.find(b => b.ID === branch.Id);
          if (branchOption) {
            return branchOption.BranchNameENG || branchOption.Title || `Branch ${branch.Id}`;
          }
        }
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

      // Enhanced picture URL handling
      const getPictureUrl = (pictureValue: any): string => {
        if (!pictureValue) return '';
        
        console.log("🔍 Processing picture value:", pictureValue);
        
        // If it's already a full URL
        if (typeof pictureValue === 'string' && pictureValue.startsWith('http')) {
          return pictureValue;
        }
        
        // If it's an object with URL property (common in SharePoint)
        if (typeof pictureValue === 'object' && pictureValue.Url) {
          return pictureValue.Url;
        }
        
        // If it's a server-relative URL, convert to absolute
        if (typeof pictureValue === 'string' && pictureValue.startsWith('/')) {
          const baseUrl = this.context.pageContext.web.absoluteUrl;
          return `${baseUrl}${pictureValue}`;
        }
        
        // If it's just a path, try to construct URL
        if (typeof pictureValue === 'string') {
          const baseUrl = this.context.pageContext.web.absoluteUrl;
          return `${baseUrl}/${pictureValue}`.replace(/([^:]\/)\/+/g, '$1');
        }
        
        return '';
      };

      const pictureUrl = getPictureUrl(emp.Picture);
      console.log("🖼️ Final picture URL:", pictureUrl);
      
      // Create HTML for the view-only profile with centered header
      let formHTML = `
        <div class="profile-container">
          <!-- Header Section with Centered Content -->
          <div class="profile-header">
            <div class="header-content-centered">
              <h1 class="profile-title">My Profile</h1>
              <div class="user-info-centered">
                <div class="user-avatar-centered">
                  ${pictureUrl ? `
                    <img src="${pictureUrl}" alt="Profile Picture" class="avatar-img-centered" onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';" />
                    <div class="avatar-placeholder-centered" style="display: none;">👤</div>
                  ` : `
                    <div class="avatar-placeholder-centered">👤</div>
                  `}
                </div>
                <div class="user-details-centered">
                  <h2 class="user-name-centered">${displayValue(emp.Name_x002d_EN)}</h2>
                  <p class="user-code-centered">Employee Code: ${displayValue(emp.Code)}</p>
                  <p class="user-email-centered">${currentUser.Email}</p>
                </div>
              </div>
            </div>
          </div>

          <!-- Main Content - All Fields Read Only -->
          <div class="profile-content">
            <!-- Professional Information Section -->
            <section class="section professional-section">
              <div class="section-header">
                <div class="section-icon">💼</div>
                <h3>Professional Information</h3>
              </div>
              
              <div class="fields-grid">
                <div class="field-group">
                  <label>Department</label>
                  <div class="field-value">${displayValue(emp.Department)}</div>
                </div>
                <div class="field-group">
                  <label>Specialty</label>
                  <div class="field-value">${displayValue(emp.Specialty)}</div>
                </div>
                <div class="field-group">
                  <label>Sub Speciality</label>
                  <div class="field-value">${displayValue(emp.SubSpeciality)}</div>
                </div>
                <div class="field-group">
                  <label>Degree</label>
                  <div class="field-value">${displayValue(emp.Degree)}</div>
                </div>

                <div class="field-group">
                  <label>Title (AR)</label>
                  <div class="field-value">${displayValue(emp.Title_x002d_AR)}</div>
                </div>
                <div class="field-group">
                  <label>Title (EN)</label>
                  <div class="field-value">${displayValue(emp.TitleEN)}</div>
                </div>

                <div class="field-group">
                  <label>Revenue</label>
                  <div class="field-value">${displayValue(emp.Revenue)}</div>
                </div>
                <div class="field-group">
                  <label>Exclusive</label>
                  <div class="field-value">${displayValue(emp.Exclusive)}</div>
                </div>
                <div class="field-group">
                  <label>Army Status</label>
                  <div class="field-value">${displayValue(emp.ArmyStatus)}</div>
                </div>
              </div>
            </section>

            <!-- Schedule & Branch Information Section -->
            <section class="section schedule-section">
              <div class="section-header">
                <div class="section-icon">📅</div>
                <h3>Schedule & Branch Information</h3>
              </div>
              
              <div class="schedule-container">
                <div class="schedule-row">
                  <div class="field-group">
                    <label>Start Date 1</label>
                    <div class="field-value">${formatDate(emp.StartDate1)}</div>
                  </div>
                  <div class="field-group">
                    <label>Branch 1</label>
                    <div class="field-value">${getBranchDisplayName(branch1ENG)}</div>
                  </div>
                  <div class="field-group">
                    <label>Shift 1</label>
                    <div class="field-value">${displayValue(emp.Shift1)}</div>
                  </div>
                </div>

                <div class="schedule-row">
                  <div class="field-group">
                    <label>Start Date 2</label>
                    <div class="field-value">${formatDate(emp.StartDate2)}</div>
                  </div>
                  <div class="field-group">
                    <label>Branch 2</label>
                    <div class="field-value">${getBranchDisplayName(branch2ENG)}</div>
                  </div>
                  <div class="field-group">
                    <label>Shift 2</label>
                    <div class="field-value">${displayValue(emp.Shift2)}</div>
                  </div>
                </div>

                <div class="schedule-row">
                  <div class="field-group">
                    <label>Start Date 3</label>
                    <div class="field-value">${formatDate(emp.StartDate3)}</div>
                  </div>
                  <div class="field-group">
                    <label>Branch 3</label>
                    <div class="field-value">${getBranchDisplayName(branch3ENG)}</div>
                  </div>
                  <div class="field-group">
                    <label>Shift 3</label>
                    <div class="field-value">${displayValue(emp.Shift3)}</div>
                  </div>
                </div>
              </div>
            </section>

            <!-- Personal Information Section -->
            <section class="section personal-section">
              <div class="section-header">
                <div class="section-icon">👤</div>
                <h3>Personal Information</h3>
              </div>
              
              <div class="fields-grid">
                <div class="field-group">
                  <label>Name (EN)</label>
                  <div class="field-value">${displayValue(emp.Name_x002d_EN)}</div>
                </div>
                <div class="field-group">
                  <label>Name (AR)</label>
                  <div class="field-value">${displayValue(emp.Name_x002d_AR)}</div>
                </div>
                <div class="field-group">
                  <label>Email</label>
                  <div class="field-value">${displayValue(emp.Email)}</div>
                </div>

                <div class="field-group">
                  <label>Phone</label>
                  <div class="field-value">${displayValue(emp.Phone)}</div>
                </div>
                <div class="field-group">
                  <label>Date of Birth</label>
                  <div class="field-value">${formatDate(emp.DateofBirth)}</div>
                </div>
                <div class="field-group">
                  <label>Marital Status</label>
                  <div class="field-value">${displayValue(emp.MaritalStatus)}</div>
                </div>
              </div>

              <!-- Bio Fields (Full Width) -->
              <div class="bio-section">
                <div class="field-group full-width">
                  <label>Bio (AR)</label>
                  <div class="field-value bio-value">${displayValue(emp.Bio)}</div>
                </div>
                <div class="field-group full-width">
                  <label>Bio (EN)</label>
                  <div class="field-value bio-value">${displayValue(emp.BioEN)}</div>
                </div>
              </div>
            </section>

            <!-- Certifications Section -->
            <section class="section certifications-section">
              <div class="section-header">
                <div class="section-icon">📎</div>
                <h3>Certifications & Documents</h3>
              </div>
              
              <div class="attachments-container">
                <div class="attachments-list" id="attachmentsList">
      `;
      
      if (attachments.length > 0) {
        attachments.forEach(attachment => {
          formHTML += `
            <div class="attachment-item">
              <div class="attachment-info">
                <span class="attachment-icon">📄</span>
                <a href="${attachment.ServerRelativeUrl}" target="_blank" class="attachment-link">
                  ${attachment.FileName}
                </a>
              </div>
            </div>
          `;
        });
      } else {
        formHTML += `<div class="no-attachments">No certification files uploaded</div>`;
      }
      
      formHTML += `
                </div>
              </div>
            </section>
          </div>
        </div>

        <style>
          .profile-container {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 1200px;
            margin: 0 auto;
            background: #ffffff;
            min-height: 100vh;
            color: #2c3e50;
          }

          .profile-header {
            background: #ffffff;
            color: #2c3e50;
            padding: 3rem 0;
            border-bottom: 1px solid #e9ecef;
          }

          .header-content-centered {
            max-width: 1200px;
            margin: 0 auto;
            padding: 0 2rem;
            text-align: center;
          }

          .profile-title {
            font-size: 2.5rem;
            font-weight: 300;
            margin-bottom: 2rem;
            color: #2c3e50;
          }

          .user-info-centered {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 2rem;
          }

          .user-avatar-centered {
            width: 220px;
            height: 220px;
            border-radius: 50%;
            overflow: hidden;
            border: 3px solid #e9ecef;
            background: #f8f9fa;
            display: flex;
            align-items: center;
            justify-content: center;
          }

          .avatar-img-centered {
            width: 100%;
            height: 100%;
            object-fit: cover;
          }

          .avatar-placeholder-centered {
            width: 100%;
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 5rem;
            color: #6c757d;
          }

          .user-details-centered {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 0.5rem;
          }

          .user-name-centered {
            margin: 0;
            font-size: 2.5rem;
            font-weight: 600;
            color: #2c3e50;
            line-height: 1.2;
            text-align: center;
          }

          .user-code-centered {
            margin: 0;
            font-size: 1.3rem;
            color: #6c757d;
            font-weight: 500;
            text-align: center;
          }

          .user-email-centered {
            margin: 0;
            color: #6c757d;
            font-size: 1.1rem;
            text-align: center;
          }

          .profile-content {
            padding: 2rem;
            display: flex;
            flex-direction: column;
            gap: 2rem;
          }

          .section {
            background: white;
            border-radius: 8px;
            padding: 2rem;
            border: 1px solid #e9ecef;
          }

          .section-header {
            display: flex;
            align-items: center;
            gap: 1rem;
            margin-bottom: 2rem;
            padding-bottom: 1rem;
            border-bottom: 1px solid #e9ecef;
          }

          .section-icon {
            font-size: 1.5rem;
            color: #6c757d;
          }

          .section-header h3 {
            margin: 0;
            color: #2c3e50;
            font-size: 1.3rem;
            font-weight: 600;
          }

          .fields-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 1.5rem;
          }

          .schedule-container {
            display: flex;
            flex-direction: column;
            gap: 1.5rem;
          }

          .schedule-row {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 1.5rem;
            padding: 1.5rem;
            background: #f8f9fa;
            border-radius: 6px;
            border: 1px solid #e9ecef;
          }

          .schedule-row .field-group {
            margin: 0;
          }

          .field-group {
            display: flex;
            flex-direction: column;
          }

          .field-group.full-width {
            grid-column: 1 / -1;
          }

          .field-group label {
            font-weight: 600;
            color: #495057;
            margin-bottom: 0.5rem;
            font-size: 0.9rem;
          }

          .field-value {
            padding: 0.75rem 1rem;
            background: #ffffff;
            border: 1px solid #e9ecef;
            border-radius: 4px;
            color: #2c3e50;
            min-height: 44px;
            display: flex;
            align-items: center;
            font-weight: 400;
            font-size: 0.95rem;
          }

          .bio-value {
            min-height: 100px;
            align-items: flex-start;
            line-height: 1.5;
            white-space: pre-wrap;
          }

          .bio-section {
            margin-top: 1.5rem;
            padding-top: 1.5rem;
            border-top: 1px solid #e9ecef;
          }

          .attachments-container {
            margin-top: 1rem;
          }

          .attachments-list {
            display: flex;
            flex-direction: column;
            gap: 0.75rem;
          }

          .attachment-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 1rem;
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 4px;
            transition: all 0.2s ease;
          }

          .attachment-item:hover {
            background: #e9ecef;
          }

          .attachment-info {
            display: flex;
            align-items: center;
            gap: 0.75rem;
          }

          .attachment-icon {
            font-size: 1.25rem;
            color: #6c757d;
          }

          .attachment-link {
            color: #495057;
            text-decoration: none;
            font-weight: 400;
            font-size: 0.95rem;
          }

          .attachment-link:hover {
            color: #2c3e50;
            text-decoration: underline;
          }

          .no-attachments {
            text-align: center;
            color: #6c757d;
            font-style: italic;
            padding: 2rem;
            font-size: 1rem;
          }

          .error-message {
            background: white;
            padding: 2rem;
            border-radius: 8px;
            text-align: center;
            border: 1px solid #e9ecef;
            margin: 2rem;
          }

          .error-message h2 {
            color: #dc3545;
            margin-bottom: 1rem;
            font-size: 1.5rem;
          }

          @media (max-width: 768px) {
            .profile-content {
              padding: 1rem;
            }
            
            .section {
              padding: 1.5rem;
            }
            
            .schedule-row {
              grid-template-columns: 1fr;
              gap: 1rem;
            }
            
            .user-avatar-centered {
              width: 180px;
              height: 180px;
            }
            
            .avatar-placeholder-centered {
              font-size: 4rem;
            }
            
            .user-name-centered {
              font-size: 2rem;
            }
            
            .user-code-centered {
              font-size: 1.1rem;
            }
            
            .profile-title {
              font-size: 2rem;
            }
            
            .fields-grid {
              grid-template-columns: 1fr;
            }
          }

          @media (max-width: 480px) {
            .header-content-centered {
              padding: 0 1rem;
            }
            
            .profile-content {
              padding: 0.5rem;
            }
            
            .section {
              padding: 1rem;
            }
            
            .user-avatar-centered {
              width: 150px;
              height: 150px;
            }
            
            .avatar-placeholder-centered {
              font-size: 3rem;
            }
            
            .user-name-centered {
              font-size: 1.8rem;
            }
            
            .user-code-centered {
              font-size: 1rem;
            }
            
            .profile-title {
              font-size: 1.8rem;
            }
          }
        </style>
      `;
      
      this.domElement.innerHTML = formHTML;
      
      console.log("✅ View Profile rendered successfully");
      
    } catch (error) {
      console.error("❌ Error in render:", error);
      this.domElement.innerHTML = `
        <div class="profile-container">
          <div class="error-message">
            <h2>Error Loading Profile</h2>
            <p><strong>Message:</strong> ${error.message}</p>
            <p>Please check browser console for details and try refreshing the page.</p>
          </div>
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