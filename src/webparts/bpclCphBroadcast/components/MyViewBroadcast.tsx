import * as React from "react";
import { useEffect, useState } from "react";
import styles from "./MyViewBroadcast.module.scss";

//import "@fontsource/inter";
import "@fontsource/inter/500.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/700.css";

import { BroadcastService } from "../services/BroadcastService";

import BroadcastTable from "./BroardcastTable";
import CreateBroadcast from "./CreateBroadcast";
import BroadcastEditForm from "./BroadcastEditForm";
import PopupModal from "../services/PopupModel";
import PreviewBroadcastModal from "./PreviewBroadcastModel";
import ErrorService from "../services/ErrorService";


/* Props */
export interface IBroadcastProps {
  sp: any;
  context: any;
}

/* Taxonomy */
export interface ITaxonomyFieldValue {
  Label: string;
  TermGuid: string;
  WssId?: number;
}

/* Broadcast Form */
export interface IBroadcastForm {
  Title: string;

  Sendmail: boolean;

  DLGroupId?: { results: number[] };
  DLGroup?: any[];

  BroadcastType: ITaxonomyFieldValue | null;
 
  RelatedSiteURL?: string;

  EventDate: string;
  ExpiryDate?: string;

  AssignTo: number[];
  PendingWith: number[];
  FollowOnActionBy: number[];
  ViewBy: number[];

  AuthorId?: number;

  Thumbnail: string;
  ThumbnailCaption: string;

  AdditionalInformation1: string;
  AdditionalInformation2: string;
  AdditionalInformation3: string;

  Comments: string;

  Status?: string;
}

const MyViewBroadcast: React.FC<IBroadcastProps> = ({ sp, context }) => {
  const service = new BroadcastService(context);

  const [items, setItems] = useState<any[]>([]);
  const [status, setStatus] = useState("Draft");
  const [search, setSearch] = useState("");
  const [isLoading, setIsLoading] = useState(true);

  const [view, setView] = useState<"list" | "create" | "edit">("list");

  const [isSiteAdmin, setIsSiteAdmin] = useState(false);

  const [editId, setEditId] = useState<number | null>(null);

  const [termSetBroadcastTypesId, setTermSetBroadcastTypesId] =
    useState<string>("");

  const [showPopup, setShowPopup] = useState<boolean>(false);
  const [popupMessage, setPopupMessage] = useState<string>("");
  const [popupType, setPopupType] = useState<"success" | "danger" | "confirm">(
    "success",
  );

  const [showPreview, setShowPreview] = useState(false);
  const [previewItem, setPreviewItem] = useState<any>(null);
  const [previewAttachments, setPreviewAttachments] = useState<any[]>([]);

  const [canCreateBroadcast, setCanCreateBroadcast] = useState(false);

  const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
  const [existingAttachments, setExistingAttachments] = useState<any[]>([]);

  const [attachmentLinks, setAttachmentLinks] = useState<
    { FileName: string; ServerRelativeUrl: string }[]
  >([]);

  const [form, setForm] = useState<IBroadcastForm>({
    Title: "",
    Sendmail: false,
    BroadcastType: null,
    RelatedSiteURL: "",
    EventDate: "",
    ExpiryDate: "",
    AssignTo: [],
    PendingWith: [],
    FollowOnActionBy: [],
    ViewBy: [],
    Thumbnail: "",
    ThumbnailCaption: "",
    AdditionalInformation1: "",
    AdditionalInformation2: "",
    AdditionalInformation3: "",
    Comments: "",
    Status: "Draft",
  });

  useEffect(() => {
    // keep TS happy — attachments are used during save/submit
    if (existingAttachments.length >= 0) {
      // no-op
    }
  }, [existingAttachments]);

  useEffect(() => {
    const loadAdmin = async () => {
      const admin = await service.isCurrentUserSiteAdmin();
      setIsSiteAdmin(admin);
    };

    loadAdmin();
  }, []);
 

  /* ---------------- PREVIEW ---------------- */
  const openPreviewPopup = async (id: number, item?: any) => {
    const links = await service.getAttachmentLinks(id);
    setPreviewAttachments(links);

    setPreviewItem(item ? item : { ...form, Id: id });
    setShowPreview(true);
  };

  const openEditPreviewPopup = async () => {
    if (!editId) return;
    await openPreviewPopup(editId);
  };

  /* ---------------- LOAD BROADCAST ---------------- */
  const loadBroadcast = async () => {
    try {
              setIsLoading(true); //  START loading
 // 2️ Validate user is mapped in TeamCMList
      // const isValid = await service.isUserPartOfSbuTeam();

      // if (!isValid) {
      //   setCanCreateBroadcast(false);
      //   setItems([]);
      //   return;
      // }

      // 3️ Role-based create permission
      const canCreate = await service.canUserCreateByType("BroadCast");
      setCanCreateBroadcast(canCreate);

      // 4️ Load data (SBU filtering already handled inside service)
      const data = await service.getBroadcastForCreator(status, 500);

      setItems(data); 
    } catch (err) {
 ErrorService.logError(
    "MyViewBroadcast.loadBroadcast",
    err
  );   setPopupType("danger");

  setPopupMessage(
    ErrorService.getUserMessage(err)
  );

  setShowPopup(true);
   }finally {
    setIsLoading(false); //  END loading
  } 
  };

  useEffect(() => {
    loadBroadcast();
  }, [status]);

  /* ---------------- TERMSET ---------------- */
  useEffect(() => {
    (async () => {
      try {
        const id = await service.getTermSetId("BroadcastType");
        setTermSetBroadcastTypesId(id);
      } catch (error) {

  ErrorService.logError(
    "MyViewBroadcast.getTermSetId", 
    error
  );setPopupType("danger");

  setPopupMessage(
    ErrorService.getUserMessage(error)
  );

  setShowPopup(true);
      }
    })();
  }, []);

  const reloadAttachments = async (id?: number) => {
    const targetId = id ?? editId;
    if (!targetId) return;

    const files = await service.getAttachments(targetId); // ✔ correct ID
    setExistingAttachments(files || []);

    const links = await service.getAttachmentLinks(targetId);
    setAttachmentLinks(links);
  };

  /* ---------------- OPEN EDIT ---------------- */
  const openEdit = async (item: any) => {
    setUploadedFiles([]);
    setEditId(item.Id);

    setExistingAttachments([]);
    setAttachmentLinks([]);

    await reloadAttachments(item.Id);
    //load existing files immediately

    setForm({
      Title: item.Title || "",
      Sendmail: item.Sendmail || false,
      DLGroupId: item.DLGroupId ? { results: item.DLGroupId } : { results: [] },
      DLGroup: item.DLGroup || [],
      // BroadcastType: item.BroadcastType || null,
      /*  Save taxonomy for SharePoint */
      BroadcastType: item.BroadcastType?.TermGuid
        ? {
            Label: item.BroadcastType.Label,
            TermGuid: item.BroadcastType.TermGuid,
          }
        : null,
      RelatedSiteURL: item.RelatedSiteURL?.Url || "",
      EventDate: item.EventDate
        ? new Date(item.EventDate).toISOString().split("T")[0]
        : "",

      ExpiryDate: item.ExpiryDate
        ? new Date(item.ExpiryDate).toISOString().split("T")[0]
        : "",
      AssignTo: item.AssignToId ? [item.AssignToId] : [],
      PendingWith: item.PendingWithId ? [item.PendingWithId] : [],
      FollowOnActionBy: item.FollowOnActionById
        ? [item.FollowOnActionById]
        : [],
      ViewBy: item.ViewById ? [item.ViewById] : [],
      AuthorId: item.AuthorId || null,
      Thumbnail: item.Thumbnail || "",
      ThumbnailCaption: item.ThumbnailCaption || "",
      AdditionalInformation1: item.AdditionalInformation1 || "",
      AdditionalInformation2: item.AdditionalInformation2 || "",
      AdditionalInformation3: item.AdditionalInformation3 || "",
      Comments: item.Comments || "",
      Status: item.Status || "Draft",
    });

    setView("edit");
  };

  const showError = (msg: string) => {
    setPopupType("danger");
    setPopupMessage(msg);
    setShowPopup(true);
  };

/* ---------------- SAVE DRAFT ---------------- */
const saveDraft = async () => {

  try {

    if (!editId) return;

    if (!form.Title?.trim()) {

      return showError(
        "Title is required"
      );
    }

    const currentUserId =
      await service.getCurrentUserId();

    // If logged-in user is PendingWith → Approver saving
    const isApproverSaving =
      form.PendingWith?.[0] === currentUserId;

    const payload: any = {

      Title: form.Title,

      Sendmail: form.Sendmail,

      DLGroupId:
        form.DLGroupId?.results || [],

      BroadcastType:
        form.BroadcastType
          ? service.getTaxonomyValue(
              form.BroadcastType
            )
          : null,

      EventDate: form.EventDate
        ? new Date(
            form.EventDate
          ).toISOString()
        : null,

      RelatedSiteURL:
        form.RelatedSiteURL
          ? {
              Url: form.RelatedSiteURL,
              Description:
                form.RelatedSiteURL,
            }
          : null,

      AssignToId:
        form.AssignTo?.[0] || null,

      PendingWithId:
        isApproverSaving
          ? form.PendingWith?.[0] || null
          : form.AssignTo?.[0] || null,

      FollowOnActionById:
        form.FollowOnActionBy?.[0] || null,

      ViewById:
        form.ViewBy?.[0] || null,

      ThumbnailCaption:
        form.ThumbnailCaption,

      AdditionalInformation1:
        form.AdditionalInformation1,

      AdditionalInformation2:
        form.AdditionalInformation2,

      AdditionalInformation3:
        form.AdditionalInformation3,

      Comments: form.Comments,
    };

    // Only Creator should change to Draft
    if (!isApproverSaving) {

      payload.Status = "Draft";
    }

    // SAVE EVENT
    await service.updateBroadcast(
      editId,
      payload
    );

    // Upload new attachments
    if (
      uploadedFiles.length > 0 &&
      editId
    ) {

      await service.uploadAttachments(
        editId,
        uploadedFiles
      );

      setUploadedFiles([]);

      // wait SharePoint
      await new Promise((r) =>
        setTimeout(r, 400)
      );

      await reloadAttachments();
    }

    /* Popup */
    const formatRequestNumber = (
      id: number
    ) => {

      if (!id) return "-";

      const padded =
        ("00000000" + id.toString())
          .slice(-8);

      return `RQ${padded}`;
    };

    setPopupType("success");

    setPopupMessage(
      `Request Number: ${formatRequestNumber(editId)} saved successfully.`
    );

    setShowPopup(true);

    await loadBroadcast();

    setView("list");

  } catch (error) {

    ErrorService.logError(
      "BroadcastEditForm.saveDraft",
      error
    );

    setPopupType("danger");

    setPopupMessage(
      ErrorService.getUserMessage(error)
    );

    setShowPopup(true);
  }
};

 /* ---------------- SUBMIT EVENT ---------------- */
const submitBroadcast = async () => {

  try {

    if (!editId) return;

    /* ---------- REQUIRED FIELD VALIDATION ---------- */

    if (!form.Title?.trim()) {

      return showError(
        "Title is required"
      );
    }

    if (form.Title.length > 255) {

      return showError(
        "Title should not exceed 255 characters."
      );
    }

    if (!form.DLGroupId?.results?.length) {

      return showError(
        "DL Group is required"
      );
    }

    if (!form.BroadcastType) {

      return showError(
        "Broadcast Type is required"
      );
    }

    if (!form.EventDate) {

      return showError(
        "Event Date is required"
      );
    }

    if (!form.AssignTo?.length) {

      return showError(
        "Approver is required"
      );
    }

    if (!form.FollowOnActionBy?.length) {

      return showError(
        "Content Manager is required"
      );
    }

    /* ---------- RELATED DOCUMENT REQUIRED ---------- */

    if (
      (!uploadedFiles ||
        uploadedFiles.length === 0) &&
      (!existingAttachments ||
        existingAttachments.length === 0)
    ) {

      return showError(
        "Please upload at least one related document"
      );
    }

    if (!form.ThumbnailCaption?.trim()) {

      return showError(
        "Image Description is required"
      );
    }

    await service.updateBroadcast(editId, {

      Title: form.Title,

      Sendmail: form.Sendmail,

      DLGroupId:
        form.DLGroupId?.results || [],

      BroadcastType:
        form.BroadcastType
          ? service.getTaxonomyValue(
              form.BroadcastType
            )
          : null,

      EventDate: form.EventDate
        ? new Date(
            form.EventDate
          ).toISOString()
        : null,

      RelatedSiteURL:
        form.RelatedSiteURL
          ? {
              Url: form.RelatedSiteURL,
              Description:
                form.RelatedSiteURL,
            }
          : null,

      AssignToId:
        form.AssignTo?.[0] || null,

      PendingWithId:
        form.AssignTo?.[0] || null,

      FollowOnActionById:
        form.FollowOnActionBy?.[0] || null,

      ViewById:
        form.ViewBy?.[0] || null,

      ThumbnailCaption:
        form.ThumbnailCaption,

      AdditionalInformation1:
        form.AdditionalInformation1,

      AdditionalInformation2:
        form.AdditionalInformation2,

      AdditionalInformation3:
        form.AdditionalInformation3,

      Comments: form.Comments,

      Status: "PendingWithLM",
    });

    // Upload new attachments
    if (
      uploadedFiles.length > 0 &&
      editId
    ) {

      await service.uploadAttachments(
        editId,
        uploadedFiles
      );

      setUploadedFiles([]);

      // wait SharePoint
      await new Promise((r) =>
        setTimeout(r, 400)
      );

      await reloadAttachments();
    }

    // wait for SharePoint calculated column to update
    await new Promise((r) =>
      setTimeout(r, 800)
    );

    const expiry =
      await service.getExpiryDateById(
        editId
      );

    if (expiry) {

      setForm((prev) => ({
        ...prev,
        ExpiryDate:
          expiry.split("T")[0],
      }));
    }

    const formatRequestNumber = (
      id: number
    ) => {

      if (!id) return "-";

      const padded =
        ("00000000" + id.toString())
          .slice(-8);

      return `RQ${padded}`;
    };

    setPopupType("success");

    setPopupMessage(
      `Request Number: ${formatRequestNumber(editId)} submitted successfully.`
    );

    setShowPopup(true);

    await loadBroadcast();

    setView("list");

  } catch (error) {

    ErrorService.logError(
      "BroadcastEditForm.submitBroadcast",
      error
    );

    setPopupType("danger");

    setPopupMessage(
      ErrorService.getUserMessage(error)
    );

    setShowPopup(true);
  }
}; 

  return (
    <div className={styles.wrapper}> 
      {view === "list" && (
        <BroadcastTable
          context={context}
          items={items}
                      isLoading={isLoading}

          selectedStatus={status}
          isSiteAdmin={isSiteAdmin}
          search={search}
          onStatusChange={setStatus}
          onSearchChange={setSearch}
          onEdit={openEdit}
          onCreate={() => setView("create")}
          canCreateBroadcast={canCreateBroadcast}
          onPreviewPopup={(item: any) => openPreviewPopup(item.Id, item)}
        />
      )}

      {view === "create" && (
        <CreateBroadcast
          service={service}
          context={context}
          //form={form}
          //setForm={setForm}
          onBack={() => setView("list")}
          termSetBroadcastTypesId={termSetBroadcastTypesId}
          onEditAfterSave={async (newId: number) => {
            await loadBroadcast();

            const latest = await service.getBroadcastForCreator("All", 500);
            const createdItem = latest.find((x: any) => x.Id === newId);

            if (createdItem) {
              await openEdit(createdItem);
            } else {
              setEditId(newId);
              setView("edit");
            }
          }}
        />
      )}

      {view === "edit" && (
        <BroadcastEditForm
          service={service}
          context={context}
          form={form}
          setForm={setForm}
          editId={editId}
          onBack={async () => {
            await loadBroadcast();
            setView("list");
          }}
          refreshList={loadBroadcast}
          uploadedFiles={uploadedFiles}
          setUploadedFiles={setUploadedFiles}
          existingAttachments={existingAttachments}
          reloadAttachments={reloadAttachments}
          onPreviewPopup={openEditPreviewPopup}
          onSaveDraft={saveDraft}
          onSubmit={submitBroadcast}
          termSetBroadcastTypesId={termSetBroadcastTypesId}
          attachmentLinks={attachmentLinks}
        />
      )}

      <PreviewBroadcastModal
        show={showPreview}
        onClose={() => setShowPreview(false)}
        item={previewItem}
        attachments={previewAttachments}
      />

      <PopupModal
        show={showPopup}
        type={popupType}
        message={popupMessage}
        onClose={() => setShowPopup(false)}
      />
    </div>
  );
};

export default MyViewBroadcast;
