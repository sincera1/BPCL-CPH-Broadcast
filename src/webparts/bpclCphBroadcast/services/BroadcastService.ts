import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import "@pnp/sp/attachments";
import "@pnp/sp/profiles";
import { SPHttpClient } from "@microsoft/sp-http";

import AuditService from "../services/AuditService";
import ErrorService from "./ErrorService";


export class BroadcastService {
  public sp;
  private context: any;

  private auditService: AuditService;

  constructor(context: any) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));

    this.auditService = new AuditService(
      this.sp,
      context.pageContext.web.absoluteUrl,
    );
  }

  //geeting site colloetion admin
  public async isCurrentUserSiteAdmin(): Promise<boolean> {
    try {
      const user = await this.sp.web.currentUser();

      // IsSiteAdmin = true only for Site Collection Admin
      return user?.IsSiteAdmin === true;
    } catch (error) {
  ErrorService.logError("BroadcastService.isCurrentUserSiteAdmin", error);
      return false;
    }
  }

  public async getAllCounts(): Promise<{
    news: number;
    events: number;
    broadcasts: number;
    brands: number;
  }> {
    try {
      let news = 0;
      let events = 0;
      let broadcasts = 0;
      let brands = 0;

      const { profile, sbu } = await this.getUserContext();

      let filter = "";

      if (sbu) {
        let teamName = (sbu || "").trim();

        if (
          teamName === "HRS" ||
          teamName === "HRD" ||
          teamName === "HUMAN RESOURCES"
        ) {
          filter = "(SBU eq 'HRS' or SBU eq 'HRD' or SBU eq 'HUMAN RESOURCES')";
        } else {
          filter = `SBU eq '${teamName.replace(/'/g, "''")}'`;
        }
      } else if (profile.department) {
        filter = `Department eq '${profile.department.replace(/'/g, "''")}'`;
      }

      if (!filter) {
        return { news: 0, events: 0, broadcasts: 0, brands: 0 };
      }

      const items = await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.select("CommunicationType")
        .filter(`${filter} and Status eq 'Published'`)
        .top(5000)(); //  CORRECT

      items.forEach((item: any) => {
        const type = item.CommunicationType;

        if (type === "News") news++;
        else if (type === "Event") events++;
        else if (type === "BroadCast") broadcasts++;
        else if (type === "Brand") brands++;
      });

      return { news, events, broadcasts, brands };
    } catch (error) {
      return { news: 0, events: 0, broadcasts: 0, brands: 0 };
    }
  }

 public async getLoggedInUserProfileInfo(): Promise<{
  department: string;
  reportingManagerLogin: string;
  reportingManagerUser?: any;
}> {
  try {

    const profile = await this.sp.profiles.myProperties();

    const dept =
      profile.UserProfileProperties?.find(
        (p: any) =>
          p.Key === "Department" ||
          p.Key === "SPS-Department",
      )?.Value || "";

    const managerLogin =
      profile.UserProfileProperties?.find(
        (p: any) =>
          p.Key === "Manager" ||
          p.Key === "SPS-Manager",
      )?.Value || "";

    let managerUser = null;

    if (managerLogin) { 
      try {

        managerUser = await this.sp.web.ensureUser(managerLogin);

      } catch (error) {

        ErrorService.logError(
          "BroadcastService.ensureUser",
          error
        );
      }
    }

    return {
      department: dept,
      reportingManagerLogin: managerLogin,
      reportingManagerUser: managerUser,
    };

  } catch (error) {

    ErrorService.logError(
      "BroadcastService.getLoggedInUserProfileInfo",
      error
    );

    throw new Error(
      ErrorService.getUserMessage(error)
    );
  }
}


  private async getUserGroups(): Promise<Set<string>> {
    const client = await this.context.msGraphClientFactory.getClient("3");

    let requestUrl: string | null =
      "/me/transitiveMemberOf?$select=displayName";

    const userGroups = new Set<string>();

    while (requestUrl) {
      const graphResponse: any = await client.api(requestUrl).get();

      if (graphResponse.value) {
        graphResponse.value.forEach((g: { displayName?: string }) => {
          if (g.displayName) {
            userGroups.add(g.displayName.trim().toLowerCase());
          }
        });
      }

      requestUrl = graphResponse["@odata.nextLink"]
        ? graphResponse["@odata.nextLink"].replace(
            "https://graph.microsoft.com/v1.0",
            "",
          )
        : null;
    }

    //     alert(
    //   "Groups = " +
    //   Array.from(userGroups).join(", ")
    // );

    return userGroups;
  }

  public async getCurrentUserSBUFromGraph(): Promise<string> {
    try {
      const client = await this.context.msGraphClientFactory.getClient("3");

      const user = await client
        .api("/me?$select=onPremisesExtensionAttributes")
        .get();

      return (
        user?.onPremisesExtensionAttributes?.extensionAttribute4?.trim() || ""
      );
    } catch (error) {
  ErrorService.logError(
    "BroadcastService.getCurrentUserSBUFromGraph",
    error
  );
      return "";
    }
  }

  // private buildTeamFilter(profile: any, sbu: string): string {
  //   let teamName = (sbu || "").trim();
  //   let department = (profile?.department || "").trim();

  //   // Normalize HR
  //   if (
  //     teamName === "HRS" ||
  //     teamName === "HRD" ||
  //     teamName === "HUMAN RESOURCES"
  //   ) {
  //     teamName = "HUMAN RESOURCES";
  //   }

  //   const safeTeam = teamName.replace(/'/g, "''");
  //   const safeDept = department.replace(/'/g, "''");

  //   if (safeTeam && safeDept) {
  //     return `(TeamName/TeamName eq '${safeTeam}' or TeamName/TeamName eq '${safeDept}')`;
  //   } else if (safeTeam) {
  //     return `TeamName/TeamName eq '${safeTeam}'`;
  //   } else if (safeDept) {
  //     return `TeamName/TeamName eq '${safeDept}'`;
  //   }

  //   return "";
  // }
  private buildTeamFilter(profile: any, sbu: string): string {
  const teamName = (sbu || "").trim().toUpperCase();
  const department = (profile?.department || "").trim();

  const isHR = ["HRS", "HRD", "HUMAN RESOURCES"].includes(teamName);

  const safeDept = department.replace(/'/g, "''");

  const teamFilter = isHR
    ? "TeamName/TeamName eq 'HUMAN RESOURCES'"
    : teamName
      ? `TeamName/TeamName eq '${teamName.replace(/'/g, "''")}'`
      : "";

  const departmentFilter = safeDept
    ? `TeamName/TeamName eq '${safeDept}'`
    : "";

  if (teamFilter && departmentFilter) {
    return `(${teamFilter} or ${departmentFilter})`;
  }

  return teamFilter || departmentFilter;
}

  private isSameUser(u: any, currentUser: any): boolean {
    return (
      u.Id === currentUser.Id ||
      u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase()
    );
  }

  private applyTeamPriority(
  items: any[],
  teamName: string,
  department: string,
  currentUser: any,
) {
  if (!items || items.length === 0) return [];

  const normalize = (v: string) => (v || "").trim().toLowerCase();

  // Normalize HR values
  if (
    normalize(teamName) === "hrs" ||
    normalize(teamName) === "hrd" ||
    normalize(teamName) === "human resources"
  ) {
    teamName = "HUMAN RESOURCES";
  }

  // const isSameUser = (u: any) =>
  //   u.Id === currentUser.Id ||
  //   u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase();

  // const isUserInRow = (row: any) => {
  //   const allUsers = [
  //     ...(row.PrimaryContentManager || []),
  //     ...(row.Approvers || []),
  //     ...(row.NewsCreators || []),
  //     ...(row.BroadcastCreators || []),
  //     ...(row.EventCreators || []),
  //     ...(row.BrandCreators || []),
  //   ];
  //   return allUsers.some(isSameUser);
  // };

  // 🔹 STEP 1: Try SBU match
  let matchedItems = [];

  if (teamName) {
    matchedItems = items.filter(
      (i) => normalize(i.TeamName?.TeamName) === normalize(teamName)
    );
  }

  // 🔹 STEP 2: If NO SBU match → fallback to Department
  if (matchedItems.length === 0 && department) {
    matchedItems = items.filter(
      (i) => normalize(i.TeamName?.TeamName) === normalize(department)
    );
  }

  // 🔹 STEP 3: Filter by user
  // return matchedItems.filter(isUserInRow);
      return matchedItems;

} 

  private _userContext: any = null;
  private async getUserContext() {
    if (this._userContext) return this._userContext;

    const [currentUser, profile, sbu] = await Promise.all([
      this.sp.web.currentUser(),
      this.getLoggedInUserProfileInfo(),
      this.getCurrentUserSBUFromGraph(),
    ]);

    this._userContext = {
      currentUser,
      profile,
      sbu,
    };

    return this._userContext;
  }

  //   private isTestMode = true; //  turn OFF after testing
  // private testDept = "TestDept";
  // private testSbu = ""; // force empty

  // private async getUserContext() {
  //   if (this._userContext) return this._userContext;

  //   const [currentUser, profile, sbu] = await Promise.all([
  //     this.sp.web.currentUser(),
  //     this.getLoggedInUserProfileInfo(),
  //     this.getCurrentUserSBUFromGraph(),
  //   ]);

  //   //  TEST OVERRIDE
  //   const finalSbu = this.isTestMode ? this.testSbu : sbu;
  //   const finalDept = this.isTestMode ? this.testDept : profile.department;

  //   this._userContext = {
  //     currentUser,
  //     profile: {
  //       ...profile,
  //       department: finalDept,
  //     },
  //     sbu: finalSbu,
  //   };

  //   return this._userContext;
  // }

  // Ensure USER or GROUP by loginName → returns principal ID
  public async ensurePrincipal(loginName: string): Promise<number> {
    if (!loginName) return 0;

    try {
      const ensured = await this.sp.web.ensureUser(loginName);
      return ensured?.Id || 0;
    } catch (e) {

        ErrorService.logError(
          "BroadcastService.ensureUser",
          e
        );  
              return 0;
    }
  }

  public async isUserPartOfSbuTeam(): Promise<boolean> {
    const userGroups = await this.getUserGroups();
    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "TeamName/TeamName",
        "BroadcastCreators/Id",
        // "BroadcastCreators/EMail",
                 "BroadcastCreators/Title",


        "Approvers/Id",
        // "Approvers/EMail",

        "PrimaryContentManager/Id",
        // "PrimaryContentManager/EMail",
      )
      .expand(
        "TeamName",
        "BroadcastCreators",
        "Approvers",
        "PrimaryContentManager",
      )
      .filter(filterQuery)();

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );

   
        return finalItems.some((row: any) => {
      const match = (users: any[]) =>
        users?.some((u: any) => {
          const isDirectUser =
            u.Id === currentUser.Id ||
            u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase();

          const isGroupMember =
            u.Title && userGroups.has(u.Title.trim().toLowerCase());

          return isDirectUser || isGroupMember;
        });

      return (
        match(row.BroadcastCreators) ||
        match(row.Approvers) ||
        match(row.PrimaryContentManager)
      );
    });
  }

  // getting current user role
  private async getCurrentUserRoles(): Promise<string[]> {
        const userGroups = await this.getUserGroups();
    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "TeamName/TeamName",
        "BroadcastCreators/Id",
        // "BroadcastCreators/EMail",
          "BroadcastCreators/Title",


        "Approvers/Id",
        // "Approvers/EMail",

        "PrimaryContentManager/Id",
        // "PrimaryContentManager/EMail",
      )
      .expand(
        "TeamName",
        "BroadcastCreators",
        "Approvers",
        "PrimaryContentManager",
      )
      .filter(filterQuery)();

    const roles: string[] = [];

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );

    const currentUserId = currentUser.Id;

    const isSameUser = (u: any) =>
      u.Id === currentUserId ||
      u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase();

    for (const item of finalItems) {
      // if (item.BroadcastCreators?.some((u: any) => isSameUser(u))) {
      //   roles.push("Creator");
      // }
    for (const u of item.BroadcastCreators || []) {
        const isDirectUser =
          u.Id === currentUserId ||
          u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase();

        const isGroupMember =
          u.Title && userGroups.has(u.Title.trim().toLowerCase());

        if (isDirectUser || isGroupMember) {
          roles.push("Creator");
          break;
        }
      }

      if (item.Approvers?.some((u: any) => isSameUser(u))) {
        roles.push("Approver");
      }

      if (item.PrimaryContentManager?.some((u: any) => isSameUser(u))) {
        roles.push("Content Manager");
      }
    }
    // ES5-safe unique logic
    const uniqueRoles: string[] = [];

    roles.forEach((role) => {
      if (uniqueRoles.indexOf(role) === -1) {
        uniqueRoles.push(role);
      }
    });

    return uniqueRoles.length ? uniqueRoles : ["User"];
  }

  public async saveBroadcast(
    title: string,
    dlGroupId: number[],
    SendMail: boolean,
    RelatedSiteUrl: string,
    //ConfidentialIndex: string,
    eventDate: string,
    // ArchivalPolicy: string,
    // EventToDate: string,
    addInfo1: string,
    addInfo2: string,
    addInfo3: string,
    comments: string,
    // thumbnailFile: File,
    thumbnailCaption: string,
    approverId: string,
    contentManagerId: string,
    BroadcastType: any,
    viewById: number | null,
    department: string,
    confidentialityIndex: any,
    archivalPolicy: any,
  ): Promise<{ itemId: number; requestNumber: string }> {
    try {
      //const sbuText = await this.getCurrentUserSbuText();
      const sbuText = await this.getCurrentUserSBUFromGraph();

      const result = await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.add({
          Title: title,

          DLGroupId: dlGroupId,

          Sendmail: SendMail,

          // /* Hyperlink — CORRECT */
          RelatedSiteURL: RelatedSiteUrl
            ? {
                Url: RelatedSiteUrl,
                Description: RelatedSiteUrl,
              }
            : null,

          // ConfidentialIndex: ConfidentialIndex || "",
          // ArchivalPolicy: ArchivalPolicy || "",

          EventDate: eventDate
            ? new Date(eventDate + "T00:00:00Z").toISOString()
            : null,

          // EventEndDate: EventToDate
          //   ? new Date(EventToDate + "T00:00:00Z").toISOString()
          //   : null,

          AdditionalInformation1: addInfo1 || "",
          AdditionalInformation2: addInfo2 || "",
          AdditionalInformation3: addInfo3 || "",
          Comments: comments || "",

          // Thumbnail: thumbnailFile?.name || "",
          ThumbnailCaption: thumbnailCaption || "",

          BroadcastType: BroadcastType
            ? {
                Label: BroadcastType.name,
                TermGuid: BroadcastType.key,
                WssId: -1,
              }
            : null,
          ConfidentialityIndex: confidentialityIndex
            ? this.getTaxonomyValue(confidentialityIndex)
            : null,

          ArchivalPolicy: archivalPolicy
            ? this.getTaxonomyValue(archivalPolicy)
            : null,

          /* People */
          AssignToId: approverId ? Number(approverId) : null,
          PendingWithId: approverId ? Number(approverId) : null,
          FollowOnActionById: contentManagerId
            ? Number(contentManagerId)
            : null,
          ViewById: viewById ?? null,

          Department: department || "",

          CommunicationType: "BroadCast",
          SBU: sbuText,
          Status: "Draft",
          PublishIn: "DL",
          PublishedDate: new Date().toISOString(),
        });

      const itemId = result.Id;
      if (!itemId) throw new Error("Item created but ID missing");

      /* ===== Request Number ===== */
      const requestNumber = `RQ${itemId.toString().padStart(8, "0")}`;

      await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.getById(itemId)
        .update({
          RequestNumber: requestNumber,
        });

      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");

      await this.auditService.log({
        action: "Create",
        operation: "POST",
        statusCode: 200,
        listName: "CorpCommunication",
        itemId: itemId,
        userRole: roleString,
      });

      return { itemId, requestNumber };
     } catch (error) {

  ErrorService.logError(
    "BroadcastService.saveBroadcast",
    error
  );

      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");

      //  FAILURE LOG
      await this.auditService.log({
        action: "Create Failed",
        operation: "POST",
        statusCode: 500,
        listName: "CorpCommunication",
        userRole: roleString,
      }); 

       throw new Error(
  ErrorService.getUserMessage(error)
);
    }
  }

  public async getExpiryDateById(id: number): Promise<string | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.getById(id)
        .select("ExpiryDate")();

      return item?.ExpiryDate || null;
    } catch (error) {
ErrorService.logError(
    "BroadcastService.getExpiryDateById",
    error
  );      return null;
    }
  }

   public async getApproversByTeam(): Promise<any[]> {
    const { profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "TeamName/TeamName",
        "Approvers/Id",
        "Approvers/Title",
      )
      .expand("TeamName", "Approvers")
      .filter(filterQuery)
      .top(500)();

    // const normalize = (v: string) => (v || "").trim().toLowerCase();

    // let finalItems = items.filter(
    //   (i: any) => normalize(i.TeamName?.TeamName) === normalize(sbu),
    // );

    // if (finalItems.length === 0) {
    //   finalItems = items.filter(
    //     (i: any) =>
    //       normalize(i.TeamName?.TeamName) === normalize(profile.department),
    //   );
    // }
    
const normalize = (value: string) => {
  const normalized = (value || "").trim().toUpperCase();

  return ["HRS", "HRD", "HUMAN RESOURCES"].includes(normalized)
    ? "HUMAN RESOURCES"
    : normalized;
};

let finalItems = items.filter(
  (item: any) =>
    normalize(item.TeamName?.TeamName) === normalize(sbu)
);

if (finalItems.length === 0) {
  finalItems = items.filter(
    (item: any) =>
      normalize(item.TeamName?.TeamName) ===
      normalize(profile.department)
  );
}


    const users: any[] = [];

    finalItems.forEach((i: any) => {
      if (i.Approvers) {
        if (Array.isArray(i.Approvers)) {
          i.Approvers.forEach((u: any) => users.push(u));
        } else {
          users.push(i.Approvers);
        }
      }
    });

    return users;
  }

  public async getContentManagersByTeam(): Promise<any[]> {
    const { profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "TeamName/TeamName",
        "PrimaryContentManager/Id",
        "PrimaryContentManager/Title",
      )
      .expand("TeamName", "PrimaryContentManager")
      .filter(filterQuery)
      .top(500)();

    // const normalize = (v: string) => (v || "").trim().toLowerCase();

    // let finalItems = items.filter(
    //   (i: any) => normalize(i.TeamName?.TeamName) === normalize(sbu),
    // );

    // if (finalItems.length === 0) {
    //   finalItems = items.filter(
    //     (i: any) =>
    //       normalize(i.TeamName?.TeamName) === normalize(profile.department),
    //   );
    // }
    
const normalize = (value: string) => {
  const normalized = (value || "").trim().toUpperCase();

  return ["HRS", "HRD", "HUMAN RESOURCES"].includes(normalized)
    ? "HUMAN RESOURCES"
    : normalized;
};

let finalItems = items.filter(
  (item: any) =>
    normalize(item.TeamName?.TeamName) === normalize(sbu)
);

if (finalItems.length === 0) {
  finalItems = items.filter(
    (item: any) =>
      normalize(item.TeamName?.TeamName) ===
      normalize(profile.department)
  );
}


    const users: any[] = [];

    finalItems.forEach((i: any) => {
      if (i.PrimaryContentManager) {
        if (Array.isArray(i.PrimaryContentManager)) {
          i.PrimaryContentManager.forEach((u: any) => users.push(u));
        } else {
          users.push(i.PrimaryContentManager);
        }
      }
    });

    return users;
  } 

  //  current user id
  // public async getCurrentUserId(): Promise<number> {
  //   const me = await this.sp.web.currentUser();
  //   return me.Id;
  // }
   public async getCurrentUserId(): Promise<number> {
  const { currentUser } = await this.getUserContext();
  return currentUser.Id;
}

  private normalizeUsers(field: any): any[] {
    if (!field) return [];
    return Array.isArray(field) ? field : [field];
  }

  // is role check creator or not
  public async isCurrentUserBroadcastCreator(): Promise<boolean> {
    const userId = await this.getCurrentUserId();
        const userGroups = await this.getUserGroups();
    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "TeamName/TeamName",
        "BroadcastCreators/Id",
        // "BroadcastCreators/EMail",
          "BroadcastCreators/Title",

        "Approvers/Id",
        // "Approvers/EMail",

        "PrimaryContentManager/Id",
        // "PrimaryContentManager/EMail",
      )
      .expand(
        "TeamName",
        "BroadcastCreators",
        "Approvers",
        "PrimaryContentManager",
      )
      .filter(filterQuery)();

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );

    return finalItems.some((row: any) => {
      // const isCreator = row.BroadcastCreators?.some(
      //   (u: any) =>
      //     u.Id === userId ||
      //     u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase(),
      // );
       
      let isCreator = false;
      for (const u of row.BroadcastCreators || []) {
        const isDirectUser =
          u.Id === userId ||
          u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase();

        const isGroupMember =
          u.Title && userGroups.has(u.Title.trim().toLowerCase());

        if (isDirectUser || isGroupMember) {
          isCreator = true;
          break;
        }
      }

      const isPrimaryCM = row.PrimaryContentManager?.some(
        (u: any) =>
          u.Id === userId ||
          u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase(),
      );

      const isApprover = row.Approvers?.some(
        (u: any) =>
          u.Id === userId ||
          u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase(),
      );

      return isCreator || isPrimaryCM || isApprover;
    });
  }

  // create button visiblity
  public async canUserCreateByType(
    type: "News" | "BroadCast" | "Event" | "Brand",
  ): Promise<boolean> {
        const userGroups = await this.getUserGroups();

    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "IsNewsCreator",
        "IsBroadcastCreator",
        "IsEventCreator",
        "IsBrandCreator",
        "NewsCreators/Id",
        // "NewsCreators/EMail",
         "NewsCreators/Title",
        "BroadcastCreators/Title",
        "EventCreators/Title",
        "BrandCreators/Title",


        "Approvers/Id",
        // "Approvers/EMail",

        "PrimaryContentManager/Id",
        // "PrimaryContentManager/EMail",
        "BroadcastCreators/Id",
        // "BroadcastCreators/EMail",
        "EventCreators/Id",
        // "EventCreators/EMail",
        "BrandCreators/Id",
        // "BrandCreators/EMail",
        "TeamName/TeamName",
      )
      .expand(
        "PrimaryContentManager",
        "Approvers",
        "NewsCreators",
        "BroadcastCreators",
        "EventCreators",
        "BrandCreators",
        "TeamName",
      )
      .filter(filterQuery)();

    if (!items.length) return false;

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );

    if (!finalItems.length) return false;

    for (const item of finalItems) {
      //  Check if type enabled
      const isTypeEnabled =
        (type === "News" && item.IsNewsCreator === true) ||
        (type === "BroadCast" && item.IsBroadcastCreator === true) ||
        (type === "Event" && item.IsEventCreator === true) ||
        (type === "Brand" && item.IsBrandCreator === true);

      if (!isTypeEnabled) continue;

      const primaryCM = this.normalizeUsers(item.PrimaryContentManager);
      const approvers = this.normalizeUsers(item.Approvers);

      const creators =
        type === "News"
          ? this.normalizeUsers(item.NewsCreators)
          : type === "BroadCast"
            ? this.normalizeUsers(item.BroadcastCreators)
            : type === "Event"
              ? this.normalizeUsers(item.EventCreators)
              : this.normalizeUsers(item.BrandCreators);

      const isPrimaryOrApprover =
        primaryCM.some((u) => this.isSameUser(u, currentUser)) ||
        approvers.some((u) => this.isSameUser(u, currentUser));

      if (isPrimaryOrApprover) return true;

      // const isExplicitCreator = creators.some((u) =>
      //   this.isSameUser(u, currentUser),
      // );
      
      let isExplicitCreator = false;

      for (const u of creators) {
        // Direct user
        if (
          u.Id === currentUser.Id ||
          u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase()
        ) {
          isExplicitCreator = true;
          break;
        }

        // Group match
        if (u.Title && userGroups.has(u.Title.trim().toLowerCase())) {
          isExplicitCreator = true;
          break;
        }
      }


      if (isExplicitCreator) return true;
    }

    return false;
  }

  //  Get Events list data
  public async getBroadcastForCreator(
    status: string,
    pageSize: number,
  ): Promise<any[]> {
    const { currentUser, profile, sbu } = await this.getUserContext();
    const userId = currentUser.Id;

    let filterQuery = `Created ge datetime'2024-12-31T00:00:00Z' and CommunicationType eq 'BroadCast'`;

    let teamName = (sbu || "").trim();

    if (
      teamName === "HRS" ||
      teamName === "HRD" ||
      teamName === "HUMAN RESOURCES"
    ) {
      teamName = "HUMAN RESOURCES";
    }

    if (teamName === "HUMAN RESOURCES") {
      filterQuery += ` and (SBU eq 'HRS' or SBU eq 'HRD' or SBU eq 'HUMAN RESOURCES')`;
    } else if (teamName) {
      filterQuery += ` and SBU eq '${teamName.replace(/'/g, "''")}'`;
    } else if (profile.department) {
      filterQuery += ` and Department eq '${profile.department.replace(/'/g, "''")}'`;
    }

    if (status !== "All") {
      if (status === "Draft") {
        filterQuery += `
      and Status eq 'Draft'
      and AuthorId eq ${userId}
    `;
      } else if (status === "Pending") {
        filterQuery += `
      and PendingWithId eq ${userId}
      and (
        Status eq 'PendingWithLM'
        or Status eq 'PendingWithCM'
        or Status eq 'PendingWithCreator'
      )
    `;
      } else if (status === "Approval Queue") {
        filterQuery += `
      and AuthorId eq ${userId}
      and (
        Status eq 'PendingWithLM'
        or Status eq 'PendingWithCM'
      )
    `;
      } else if (status === "Published") {
        filterQuery += `
      and Status eq 'Published'
      and (
        AuthorId eq ${userId}
        or AssignToId eq ${userId}
        or FollowOnActionById eq ${userId}
        or PendingWithId eq ${userId}
      )
    `;
      } else if (status === "View Requests") {
        filterQuery += `
    and (
      (
        AuthorId eq ${userId} and (
          Status eq 'Draft'
          or Status eq 'PendingWithLM'
          or Status eq 'PendingWithCM'
          or Status eq 'PendingWithCreator'
          or Status eq 'Published'
          or Status eq 'Rejected'
        )
      )
      or
      (
        AssignToId eq ${userId} and (
          Status eq 'PendingWithLM'
          or Status eq 'PendingWithCM'
          or Status eq 'Published'
          or Status eq 'Rejected'
          or Status eq 'PendingWithCreator'

        )
      )
      or
      (
        FollowOnActionById eq ${userId} and (
          Status eq 'PendingWithCM'
          or Status eq 'Published'
          or Status eq 'Rejected'
          or Status eq 'PendingWithLM'

        )
      )
      or
      (
        PendingWithId eq ${userId} and (
          Status eq 'PendingWithLM'
          or Status eq 'PendingWithCM'
          or Status eq 'Published'
          or Status eq 'Rejected'
          or Status eq 'PendingWithCreator'

        )
      ) 
    )
  `;
      }
    }

    const items = await this.sp.web.lists
      .getByTitle("CorpCommunication")
      .items.select(
        "Id",
        "Title",
        "SBU",
        "CommunicationType",
        "DLGroupId",
        "DLGroup/Id",
        "DLGroup/Title",
        "Sendmail",
        "EventDate",
        "ExpiryDate",
        "RelatedSiteURL",
        "BroadcastType",
        "BroadcastType/TermGuid",
        "BroadcastType/Label",
        "AssignToId",
        "PendingWithId",

        "PendingWith/Title",
        "FollowOnActionById",
        "ViewById",

        "Thumbnail",
        "ThumbnailCaption",

        "AdditionalInformation1",
        "AdditionalInformation2",
        "AdditionalInformation3",

        "Comments",

        "Created",
        "PublishedDate",
        "IsReassigned",

        "Status",

        "AuthorId",
        "Author/Title",
        "Attachments",
        "AttachmentFiles/FileName",
        "AttachmentFiles/ServerRelativeUrl",
      )
      .expand("Author", "DLGroup", "PendingWith", "AttachmentFiles")
      .filter(filterQuery)
      .orderBy("Created", false)
      .top(pageSize)();

    return items.map((it: any) => {
      const files = it.AttachmentFiles || [];

      const thumbFile = files.find((f: any) => f.FileName === it.Thumbnail);

      return {
        ...it,
        ThumbnailUrl: thumbFile ? thumbFile.ServerRelativeUrl : "",
        AttachmentsData: files,
      };
    });
  }

  // role check
  public async isCurrentUserApprover(): Promise<boolean> {
    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "Approvers/Id",
        // "Approvers/EMail",
        "TeamName/TeamName",
      )
      .expand("Approvers", "TeamName")
      .filter(filterQuery)();

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );

    return finalItems.some((row: any) =>
      row.Approvers?.some((u: any) => this.isSameUser(u, currentUser)),
    );
  }

  public async isCurrentUserPrimaryContentManager(): Promise<boolean> {
    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "PrimaryContentManager/Id",
        // "PrimaryContentManager/EMail",
        "TeamName/TeamName",
      )
      .expand("PrimaryContentManager", "TeamName")
      .filter(filterQuery)();

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );

    return finalItems.some((row: any) =>
      row.PrimaryContentManager?.some((u: any) =>
        this.isSameUser(u, currentUser),
      ),
    );
  }

  // deleteing permanently
  public async deleteEventPermanently(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.getById(id)
        .delete();

      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");

      //  SUCCESS
      await this.auditService.log({
        action: "Delete",
        operation: "DELETE",
        statusCode: 200,
        listName: "CorpCommunication",
        itemId: id,
        userRole: roleString,
      });
    } catch (error) {
      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");
      //  FAILURE
      await this.auditService.log({
        action: "Delete Failed",
        operation: "DELETE",
        statusCode: 500,
        listName: "CorpCommunication",
        itemId: id,
        userRole: roleString,
      });

       throw new Error(
  ErrorService.getUserMessage(error)
);
    }
  }

  //  Update
  public async updateBroadcast(id: number, payload: any) {
    try {
      await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.getById(id)
        .update(payload);

      //  Determine action based on Status
      let actionName = "Update";

      const status = payload?.Status;
      const reassigned = payload?.IsReassigned ?? false;

      if (status === "PendingWithLM" && reassigned === false) {
        actionName = "Submit";
      } else if (status === "Rejected") {
        actionName = "Reject";
      } else if (status === "PendingWithCreator") {
        actionName = "Reassign";
      } else if (status === "PendingWithCM") {
        actionName = "Approve";
      } else if (status === "Published") {
        actionName = "Approve";
      } else if (status === "PendingWithLM" && reassigned === true) {
        actionName = "Reassign";
      }

      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");
      //  SUCCESS LOG
      await this.auditService.log({
        action: actionName,
        operation: "POST",
        statusCode: 200,
        listName: "CorpCommunication",
        itemId: id,
        userRole: roleString,
      });
    } catch (error) {

  ErrorService.logError(
    "BroadcastService.updateBroadcast",
    error
  );
      //  FAILURE LOG (also dynamic)
      let failAction = "Update Failed";

      const status = payload?.Status;
      const reassigned = payload?.IsReassigned ?? false;

      if (status === "PendingWithLM" && reassigned === false) {
        failAction = "Submit Failed";
      } else if (status === "Rejected") {
        failAction = "Reject Failed";
      } else if (status === "PendingWithCreator") {
        failAction = "Reassign Failed";
      } else if (status === "PendingWithCM") {
        failAction = "Approve Failed";
      } else if (status === "PendingWithLM" && reassigned === true) {
        failAction = "Reassign Failed";
      } else if (status === "Published") {
        failAction = "Approve Failed";
      }

      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");

      await this.auditService.log({
        action: failAction,
        operation: "POST",
        statusCode: 500,
        listName: "CorpCommunication",
        itemId: id,
        userRole: roleString,
      });

       throw new Error(
  ErrorService.getUserMessage(error)
);
    }
  }
  //  get attachments
  public async getAttachmentLinks(
    itemId: number,
  ): Promise<{ FileName: string; ServerRelativeUrl: string }[]> {
    const files = await this.sp.web.lists
      .getByTitle("CorpCommunication")
      .items.getById(itemId)
      .attachmentFiles();

    return files.map((f: any) => ({
      FileName: f.FileName,
      ServerRelativeUrl: f.ServerRelativeUrl,
    }));
  }

  //  People picker helper
  public async getUserEmailById(userId: number): Promise<string> {
    const user = await this.sp.web.siteUsers.getById(userId)();
    return user?.Email || "";
  }

  //  TermSetId
  public async getTermSetId(fieldInternalName: string): Promise<string> {
    const field: any = await this.sp.web.lists
      .getByTitle("CorpCommunication")
      .fields.getByInternalNameOrTitle(fieldInternalName)();

    return field.TermSetId;
  }

  //  taxonomy helper
  public getTaxonomyValue(term: any) {
    if (!term) return null;

    return {
      Label: term.Label || term.text,
      TermGuid: term.TermGuid || term.key,
      WssId: -1,
    };
  }

  //  load terms
  public async getTermSetTerms(
    termSetId: string,
  ): Promise<{ key: string; text: string }[]> {
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/v2.1/termStore/sets/${termSetId}/terms?$select=id,labels`;

    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Failed to load terms: ${response.status} - ${text}`);
    }

    const json = await response.json();

    return (json.value || []).map((t: any) => ({
      key: t.id,
      text: t.labels?.[0]?.name || t.id,
    }));
  }
  /* =====================================================
     ATTACHMENTS
  ===================================================== */
  public async uploadAttachments(itemId: number, files: File[]): Promise<void> {
    if (!itemId || !files?.length) return;

    const item = this.sp.web.lists
      .getByTitle("CorpCommunication")
      .items.getById(itemId);

    for (const file of files) {
       // Security validation
    if (!file.type) {
      throw new Error("Invalid file.");
    }

    // Optional extra validation
    const blockedExtensions = [
      "exe",
      "bat",
      "cmd",
      "js",
      "msi",
      "dll",
      "sh",
      "php"
    ]; 
 
    const ext = file.name.split(".").pop()?.toLowerCase() || "";

    if (blockedExtensions.includes(ext)) {
      throw new Error("Unsupported file type.");
    }

      //  Read RAW binary buffer (prevents PDF corruption)
      const buffer = await file.arrayBuffer();

      await item.attachmentFiles.add(file.name, buffer);
    }
  }

  public async getAttachments(itemId: number) {
    if (!itemId) return [];

    return await this.sp.web.lists
      .getByTitle("CorpCommunication")
      .items.getById(itemId)
      .attachmentFiles()
      .then((files) =>
        files.map((f) => ({
          FileName: f.FileName,
          // cache-buster here (important)
          ServerRelativeUrl: f.ServerRelativeUrl + `?t=${Date.now()}`,
        })),
      );
  }

  public async deleteAttachment(
    itemId: number,
    fileName: string,
  ): Promise<void> {
    await this.sp.web.lists
      .getByTitle("CorpCommunication")
      .items.getById(itemId)
      .attachmentFiles.getByName(fileName)
      .delete();
  }

  public async getTaxonomyLabelByWssId(wssId: number): Promise<string> {
    if (!wssId) return "";

    const items = await this.sp.web.lists
      .getByTitle("TaxonomyHiddenList")
      .items.filter(`ID eq ${wssId}`)
      .select("Id", "Title")();

    return items.length > 0 ? items[0].Title : "";
  }
}
