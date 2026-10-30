import * as React from "react";
import { useEffect, useState, useRef } from "react";
//import "@fontsource/inter";
import "@fontsource/inter/500.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/700.css";

import { Form, Row, Col, Card, Button } from "react-bootstrap";
import styles from "./MyViewBroadcast.module.scss";
import PopupModal from "../services/PopupModel";
import PeoplePickerCustom, { Option } from "../services/PeoplePicker";
import { BroadcastService } from "../services/BroadcastService";
import ErrorService from "../services/ErrorService";


import { TaxonomyPicker } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";

interface Props {
  context: any;
  form: any;
  setForm: (data: any) => void;

  onBack: () => void;
  onSaveDraft: () => Promise<void>;
  onSubmit: () => Promise<void>;
  onPreviewPopup: () => void;

  editId: number | null;
  uploadedFiles: File[];
  setUploadedFiles: React.Dispatch<React.SetStateAction<File[]>>;
  existingAttachments: any[];
  reloadAttachments: () => Promise<void>;

  refreshList?: () => Promise<void>;

  service: BroadcastService;
  termSetBroadcastTypesId: string;

  attachmentLinks: { FileName: string; ServerRelativeUrl: string }[];
}

const BroadcastEditForm: React.FC<Props> = ({
  context,
  form,
  setForm,
  editId,
  onBack,
  refreshList,
  onSaveDraft,
  onSubmit,
  termSetBroadcastTypesId,
  service,
  onPreviewPopup,
  uploadedFiles,
  setUploadedFiles,
  existingAttachments,
  reloadAttachments,
}) => {
  /* ================= ROLE & WORKFLOW ================= */

  // dropdown data like CreateEvent
  const [approvers, setApprovers] = useState<any[]>([]);

  const [contentManagers, setContentManagers] = useState<any[]>([]);
  const [selectedContentManager, setSelectedContentManager] =
    useState<string>("");

  // ViewBy custom picker
  const [viewBy, setViewBy] = useState<Option | null>(null);
  const [dlGroups, setDlGroups] = useState<Option[]>([]);

  const [isBroadcastCreator, setIsBroadcastCreator] = useState(false);
  const [isApprover, setIsApprover] = useState(false);
  const [isPrimaryCM, setIsPrimaryCM] = useState(false);
  const [currentUserId, setCurrentUserId] = useState<number>(0);

  const [urlError, setUrlError] = useState<string>("");

  const [broadCastType, setBroadCastType] = useState<any>(null);

  const [showPopup, setShowPopup] = useState(false);
  const [popupType, setPopupType] = useState<"success" | "danger" | "confirm">(
    "success",
  );
  const [popupMessage, setPopupMessage] = useState("");
  const [confirmAction, setConfirmAction] = useState<
    (() => Promise<void>) | null
  >(null);

  /* ================= LOAD DATA ================= */

  const commentRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    const adjustHeight = (textarea: HTMLTextAreaElement | null) => {
      if (textarea) {
        textarea.style.height = "auto";
        textarea.style.height = textarea.scrollHeight + "px";
      }
    };

    adjustHeight(commentRef.current);
  }, []);

  useEffect(() => {
    const init = async () => {
      const userId = await service.getCurrentUserId();
      setCurrentUserId(userId);

      //const teamName = await service.getCurrentUserSBUFromGraph();

      setIsBroadcastCreator(await service.isCurrentUserBroadcastCreator());
      setIsApprover(await service.isCurrentUserApprover());
      setIsPrimaryCM(
        await service.isCurrentUserPrimaryContentManager(),
      );
    };

    init();
  }, []);

  const getTermNameByGuid = async (guid: string): Promise<string> => {
    try {
      const items = await service.sp.web.lists
        .getByTitle("TaxonomyHiddenList")
        .items.filter(`IdForTerm eq '${guid}'`)
        .select("Term")
        .top(1)();

      return items?.[0]?.Term || "";
   } catch (error) {
ErrorService.logError(
  "BroadcastEditForm.getTermNameByGuid",
  error
);
setPopupType("danger");

  setPopupMessage(
    ErrorService.getUserMessage(error)
  );

  setShowPopup(true); 
      return "";
    } 
  };

  useEffect(() => {
    const loadTerm = async () => {
      if (form?.BroadcastType?.TermGuid && termSetBroadcastTypesId) {
        const termName = await getTermNameByGuid(form.BroadcastType.TermGuid);

        if (!termName) return;

        setBroadCastType({
          key: form.BroadcastType.TermGuid,
          name: termName, //  REAL TERM NAME
          path: `${termSetBroadcastTypesId}|${termName}`,
          termSet: termSetBroadcastTypesId,
        });

        //  also update form so future saves use correct name
        setForm((prev: any) => ({
          ...prev,
          BroadcastType: {
            Label: termName,
            TermGuid: form.BroadcastType.TermGuid,
          },
        }));
      } else {
        setBroadCastType(null);
      }
    };

    loadTerm();
  }, [form?.BroadcastType?.TermGuid, termSetBroadcastTypesId]);

  const showError = (msg: string) => {
    setPopupType("danger");
    setPopupMessage(msg);
    setShowPopup(true);
  };
 
 const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
  if (!e.target.files) return;

  const files = Array.from(e.target.files);

  const allowedMimeTypes: { [key: string]: string[] } = {
    image: [
      "image/bmp",
      "image/gif",
      "image/png",
      "image/jpeg",
    ],
    document: [
      "application/pdf",
      "application/msword",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ],
  };

  const allowedExtensions = [
    "bmp",
    "gif",
    "png",
    "jpg",
    "jpeg",
    "pdf",
    "doc",
    "docx",
  ];

  const validFiles: File[] = [];

  for (const file of files) {
    const ext = file.name.split(".").pop()?.toLowerCase() || "";
    const mime = file.type;

    // Extension validation
    if (!allowedExtensions.includes(ext)) {
      showError(
        `${file.name} → Invalid file extension. Allowed: pdf, doc, docx, bmp, gif, png, jpg, jpeg`
      );
      e.target.value = "";
      return;
    }

    // MIME type validation
    const isValidMime =
      allowedMimeTypes.image.includes(mime) ||
      allowedMimeTypes.document.includes(mime);

    if (!isValidMime) {
      showError(`${file.name} → Invalid or unsafe file type detected.`);
      e.target.value = "";
      return;
    }

    // Block double extensions
    const invalidDoubleExtension =
      /\.(php|exe|js|bat|cmd|sh|msi|dll)$/i.test(file.name);

    if (invalidDoubleExtension) {
      showError(`${file.name} → Suspicious file detected.`);
      e.target.value = "";
      return;
    }

    // Image size
    if (allowedMimeTypes.image.includes(mime)) {
      if (file.size > 2 * 1024 * 1024) {
        showError(`${file.name} → Image should be less than 2MB`);
        e.target.value = "";
        return;
      }
    }

    // Document size
    if (allowedMimeTypes.document.includes(mime)) {
      if (file.size > 4 * 1024 * 1024) {
        showError(`${file.name} → Document should be less than 4MB`);
        e.target.value = "";
        return;
      }
    }

    validFiles.push(file);
  }

  if (validFiles.length > 0) {
    setUploadedFiles((prev) => [...prev, ...validFiles]);
  }

  e.target.value = "";
};
  const deleteExistingAttachment = async (fileName: string) => {
    if (!editId) return;
    await service.deleteAttachment(editId, fileName);
    await reloadAttachments();
  };

  // load approvers + content managers + role checks + ViewBy defaults
  useEffect(() => {
    const loadPeopleData = async () => {
      try { 
         //  Get profile (only for UI use)
      const profileInfo = await service.getLoggedInUserProfileInfo();
      const managerUser = profileInfo?.reportingManagerUser;

      //  Validate mapping (SBU + Department handled inside)
      // const isValid = await service.isUserPartOfSbuTeam();

      // if (!isValid) {
      //   return;
      // } 

        // --- Save current user id ---
        const currentUserId = await service.getCurrentUserId();
        setCurrentUserId(currentUserId);

        // --- Role checks FIRST ---
        const creator = await service.isCurrentUserBroadcastCreator();
        setIsBroadcastCreator(creator);

        // Existing team approver check
        const approverFromTeam = await service.isCurrentUserApprover();

        // approverr
        const isAssignedApprover = form.AssignTo?.[0] === currentUserId;

        setIsApprover(approverFromTeam || isAssignedApprover);

        // content manager
        const cm = await service.isCurrentUserPrimaryContentManager();
        setIsPrimaryCM(cm);
        // check role end

        // setting dropdown

        const appr = await service.getApproversByTeam();

        if (managerUser?.Id) {
          const alreadyExists = appr.some((u: any) => u.Id === managerUser.Id);
          if (!alreadyExists) {
            appr.unshift(managerUser); // add manager at top
          }
        }

        const itemAuthorId = form?.AuthorId || null;

        // Remove only creator from dropdown
        let finalApprovers = appr.filter((u: any) => u.Id !== itemAuthorId);

        //  Ensure PendingWith user exists in dropdown
        const pendingId = form?.AssignTo?.[0];

        if (pendingId) {
          const exists = finalApprovers.some((u: any) => u.Id === pendingId);

          if (!exists) {
            try {
              const pendingUser =
                await service.sp.web.siteUsers.getById(pendingId)();
              finalApprovers.unshift(pendingUser);
            } catch (error) {
ErrorService.logError(
  "BroadcastEditForm.loadPeopleData.pendingUser",
  error
);               }
          }
        }
 
        setApprovers(finalApprovers);

        // --- Load Content Managers ---
        const cms = await service.getContentManagersByTeam();
        setContentManagers(cms);

        if (form.FollowOnActionBy?.length > 0) {
          setSelectedContentManager(String(form.FollowOnActionBy[0]));
        }

        //  Load ViewBy user/group
        if (form.ViewBy?.length > 0) {
          const id = Number(form.ViewBy[0]);

          try {
            const user = await service.sp.web.siteUsers.getById(id)();

            setViewBy({
              label: `${user.Title}${user.Email ? ` (${user.Email})` : ""}`,
              value: user.Id,
              type: "User",
              title: user.Title,
              email: user.Email,
              loginName: user.LoginName,
            });
          } catch (err) {
            try {
              const group = await service.sp.web.siteGroups.getById(id)();

              setViewBy({
                label: `${group.Title} (Group)`,
                value: group.Id,
                type: "Group",
                title: group.Title,
              });
              } catch (error) {
ErrorService.logError(
  "BroadcastEditForm.loadPeopleData.viewBy",
  error
);     
              setViewBy(null);
            }
          }
        }

        //  Load DLGroup (for edit mode binding)
        if (form.DLGroup?.length) {
          setDlGroups(
            form.DLGroup.map((g: any) => ({
              id: g.Id,
              key: g.Id,
              value: g.Id,
              label: `${g.Title} (Group)`,
              type: "Group",
              title: g.Title,
              loginName: g.LoginName || g.Title,
            })),
          );
        }
     } catch (error) {
ErrorService.logError(
  "BroadcastEditForm.loadPeopleData",
  error
);

  setPopupType("danger");
  setPopupMessage(
    ErrorService.getUserMessage(error)
  );

  setShowPopup(true);

      }
    };

    loadPeopleData();
  }, [form?.AuthorId, form?.AssignTo, form?.FollowOnActionBy, form?.Status]);

  // const pendingWithId = form.PendingWith?.[0] || null;
  // const isPendingWithMe = pendingWithId === currentUserId;
   const pendingWithId = Number(form.PendingWith?.[0] || 0);
const isPendingWithMe =
  pendingWithId > 0 &&
  pendingWithId === Number(currentUserId);

  const assignedtoId = form.AssignTo?.[0] || null;

 // const FollowOnActionById = form.FollowOnActionBy?.[0] || null; 

  const currentStatus = (form.Status || "").toString().trim().toLowerCase();

  // ---------------- STATUS ----------------

  const isDraft = currentStatus === "draft";
  const isPendingCreator = currentStatus === "pendingwithcreator";
  const isPendingLM = currentStatus === "pendingwithlm";
  const isPendingCM = currentStatus === "pendingwithcm";

  const isFinalStatus =
    currentStatus === "rejected" || currentStatus === "published";

  // ---------------- CREATOR ----------------
  // const canCreatorEdit = isBroadcastCreator && (isDraft || isPendingCreator);
  const canCreatorEdit =
    isBroadcastCreator &&
    form.AuthorId === currentUserId &&
    (isDraft || isPendingCreator);

  // ---------------- LM ----------------
  //const canShowLMButtons = isPendingLM && isApprover && isPendingWithMe;
  
  // Approver: TeamCMList OR direct assignment
const canActAsApprover =
  isApprover || isPendingWithMe;

// Line Manager permissions
const canShowLMButtons =
  isPendingLM &&
  isPendingWithMe &&
  canActAsApprover;
  

  // ---------------- CM ----------------
  const canShowCMButtons = isPendingCM && isPrimaryCM && isPendingWithMe;

  const isCMOwnDraft =
    isPrimaryCM && form.AuthorId === currentUserId && isDraft;

  // ---------------- SAVE ----------------
  const canShowSaveButton =
    !isFinalStatus && (canCreatorEdit || canShowLMButtons || isCMOwnDraft);

  // ---------------- SUBMIT (ONLY CREATOR+cm) ----------------
  // const canShowSubmitButton = !isFinalStatus && canCreatorEdit;
  const canShowSubmitButton =
    !isFinalStatus && (canCreatorEdit || isCMOwnDraft);

  // ---------------- DELETE ----------------
  const canShowDeleteButton =
    !isFinalStatus && (canCreatorEdit || canShowLMButtons || isCMOwnDraft);

  // ---------------- APPROVE / REJECT ----------------
  const canShowApproveRejectButtons =
    !isFinalStatus && (canShowLMButtons || canShowCMButtons);

  // ---------------- FIELD EDITABLE ----------------
  //const canEditFields = canCreatorEdit || canShowLMButtons || canShowCMButtons;
  const canEditFields = canCreatorEdit || canShowLMButtons;

  const disableWorkflowDropdowns = !canEditFields || canShowLMButtons;

  // ---------------- COMMENT EDIT ----------------

  // CM can edit comment only when PendingWithCM and assigned to them
  const canCMEditComment =
    isPrimaryCM && isPendingCM && isPendingWithMe && !isFinalStatus;

  // Final comment permission
  const canEditComment = canEditFields || canCMEditComment;

  const formatRequestNumber = (id: number | null) => {
    if (!id) return "-";
    const padded = ("00000000" + id.toString()).slice(-8);
    return `RQ${padded}`;
  };

  const showActionPopup = (
    msg: string,
    type: "success" | "danger" = "success",
  ) => {
    setPopupType(type);
    setPopupMessage(`Request Number: ${formatRequestNumber(editId)} - ${msg}`);
    setShowPopup(true);
  };

  // Approve
  const onApprove = async () => {
    if (!editId) return;

    try {
      // LM / Approver
      // if (isApprover && isPendingLM && isPendingWithMe) {
      if (canShowLMButtons) {
        if (!editId) return;


       const approverId = form.AssignTo?.[0] || null;
      const cmId = form.FollowOnActionBy?.[0] || null;
      const isSamePerson = approverId && cmId && approverId === cmId;

        // Save ALL fields + move workflow in ONE update
        await service.updateBroadcast(editId, {
          Title: form.Title,

          Sendmail: form.Sendmail,

          DLGroupId: form.DLGroupId?.results || [],

          BroadcastType: form.BroadcastType
            ? service.getTaxonomyValue(form.BroadcastType)
            : null,

          EventDate: form.EventDate
            ? new Date(form.EventDate).toISOString()
            : null,
          RelatedSiteURL: form.RelatedSiteURL
            ? {
                Url: form.RelatedSiteURL,
                Description: form.RelatedSiteURL,
              }
            : null,

          //   EventEndDate: form.EventEndDate
          //     ? new Date(form.EventEndDate).toISOString()
          //     : null,

          AssignToId: form.AssignTo?.[0] || null,
          FollowOnActionById: form.FollowOnActionBy?.[0] || null,

          ViewById: form.ViewBy?.[0] || null,

          //Thumbnail: thumbnailFile ? thumbnailFile.name : form.Thumbnail,
          ThumbnailCaption: form.ThumbnailCaption,

          AdditionalInformation1: form.AdditionalInformation1,
          AdditionalInformation2: form.AdditionalInformation2,
          AdditionalInformation3: form.AdditionalInformation3,

          Comments: form.Comments,

          // //  Workflow Move
          // Status: "PendingWithCM",
          // PendingWithId: FollowOnActionById ? Number(FollowOnActionById) : null, // Workflow decision
        Status: isSamePerson ? "Published" : "PendingWithCM",
        PublishedDate: isSamePerson ? new Date().toISOString() : null,
        PendingWithId: isSamePerson ? null : cmId
        });

        //  Upload new attachments on Approve
        if (uploadedFiles.length > 0 && editId) {
          await service.uploadAttachments(editId, uploadedFiles);
          setUploadedFiles([]);

          await reloadAttachments();
        }

        await new Promise((r) => setTimeout(r, 800));

        const expiry = await service.getExpiryDateById(editId);

        if (expiry) {
          setForm((prev: any) => ({
            ...prev,
            ExpiryDate: expiry.split("T")[0],
          }));
        }

        // showActionPopup("Approved Successfully");
         showActionPopup(
        isSamePerson
          ? "Published Successfully"
          : "Approved Successfully"
      ); 

        return;
      }

      // CM → Publish
      if (isPrimaryCM && isPendingCM && isPendingWithMe) {
        await service.updateBroadcast(editId, {
          Status: "Published",
          PublishedDate: new Date().toISOString(),
          PendingWithId: null,
        });

        showActionPopup("Published Successfully");

        return;
      }

      showActionPopup("You are not allowed to Approve.", "danger");
     } catch (error) {
ErrorService.logError(
  "BroadcastEditForm.onApprove",
  error
);   
showActionPopup(
  ErrorService.getUserMessage(error),
  "danger"
);
    }
  };

  const onReject = async () => {
    if (!editId) return;

    if (!form.Comments || !form.Comments.trim()) {
      setPopupType("danger");
      setPopupMessage("Comment is required before this action.");
      setShowPopup(true);
      return;
    }
    setPopupType("confirm");
    setPopupMessage(
      `Request Number: ${formatRequestNumber(editId)} - Are you sure you want to Reject this Broadcast?`,
    );

    setConfirmAction(() => async () => {
      try {
        // await updateStatus("Rejected");
        await service.updateBroadcast(editId, {
          Status: "Rejected",
          PendingWithId: null,
          Comments: form.Comments,
        });

        setPopupType("success");
        setPopupMessage(
          `Request Number: ${formatRequestNumber(editId)} - Rejected Successfully.`,
        );
        setShowPopup(true);
 } catch (error) {
ErrorService.logError(
  "BroadcastEditForm.onReject",
  error
);   
        setPopupType("danger");
         setPopupMessage(
  `Request Number: ${formatRequestNumber(editId)} - ${ErrorService.getUserMessage(error)}`
);
        setShowPopup(true);
      }
    });

    setShowPopup(true);
  };

  // Delete
  const onDelete = async () => {
    if (!editId) return;

    setPopupType("confirm");
    setPopupMessage(
      `Request Number: ${formatRequestNumber(editId)} - Are you sure you want to delete this Broadcast permanently?`,
    );

    setConfirmAction(() => async () => {
      try {
        await service.deleteEventPermanently(editId);

        setPopupType("success");
        setPopupMessage(
          `Request Number: ${formatRequestNumber(editId)} - Deleted permanently.`,
        );
        setShowPopup(true);
      } catch (error) {
ErrorService.logError(
  "BroadcastEditForm.onDelete",
  error
);   
        setPopupType("danger");
        setPopupMessage(
  `Request Number: ${formatRequestNumber(editId)} - ${ErrorService.getUserMessage(error)}`
);
        setShowPopup(true);
      }
    });

    setShowPopup(true);
  };

  // ReAssign (with Confirm)
  const onReAssign = async () => {
    if (!editId) return;

    if (!form.Comments || !form.Comments.trim()) {
      setPopupType("danger");
      setPopupMessage("Comment is required before this action.");
      setShowPopup(true);
      return;
    }

    const creatorId = form.AuthorId;
    const approverId = assignedtoId;

    if (!creatorId || !approverId) {
      return;
    }

    // Show confirm popup
    setPopupType("confirm");
    setPopupMessage(
      `Request Number: ${formatRequestNumber(
        editId,
      )} - Are you sure you want to Re-Assign this Request?`,
    );

    setConfirmAction(() => async () => {
      try {
        // --- LM → Send back to Creator ---
        // if (isApprover && isPendingLM && isPendingWithMe) {
        if (canShowLMButtons) {
          await service.updateBroadcast(editId, {
            Status: "PendingWithCreator",
            PendingWithId: creatorId,
            Comments: form.Comments,
          });
        } 

        // --- CM → Send back to LM ---
        else if (isPrimaryCM && isPendingCM && isPendingWithMe) {
          await service.updateBroadcast(editId, {
            Status: "PendingWithLM",
            PendingWithId: approverId,
            Comments: form.Comments,
            IsReassigned: true,
          });
        }

        // 🔹 IMPORTANT — reset confirm state first
        setConfirmAction(null);

        // 🔹 Show success popup (this was missing before)
        setPopupType("success");
        setPopupMessage(
          `Request Number: ${formatRequestNumber(
            editId,
          )} - Re-Assigned Successfully.`,
        );
        setShowPopup(true);
       } catch (error) {
ErrorService.logError(
  "BroadcastEditForm.onReAssign",
  error
);  
        setConfirmAction(null);
        setPopupType("danger");
        setPopupMessage(
  `Request Number: ${formatRequestNumber(editId)} - ${ErrorService.getUserMessage(error)}`
);
        setShowPopup(true);
      }
    });

    setShowPopup(true);
  }; 

  /* ================= UI ================= */
  const isValidUrl = (url: string): boolean => {
    if (!url) return false;

    const pattern = new RegExp(
      "^(https?:\\/\\/)" +
        '((([a-zA-Z0-9$-_@.&+!*"(),]|(%[0-9a-fA-F][0-9a-fA-F]))+\\.)+' +
        "([a-zA-Z]{2,}))" +
        "(\\:[0-9]{1,5})?" +
        "(\\/.*)?$",
      "i",
    );

    return pattern.test(url.trim());
  };

  const getExpiryPreview = () => {
    const dt = new Date();
    dt.setDate(dt.getDate() + 90);

    const yyyy = dt.getFullYear();
    const mm = ("0" + (dt.getMonth() + 1)).slice(-2);
    const dd = ("0" + dt.getDate()).slice(-2);

    return `${yyyy}-${mm}-${dd}`;
  };

  /* ================= INPUT VALIDATION HELPER ================= */

  //const allowedRegex = /^[A-Za-z0-9()&_\-+=\[\]{}:;'",.?/₹$!@%*\s]*$/;
  const allowedRegex = /^[A-Za-z0-9\u0900-\u097F()&_\-+=\[\]{}:;'",.?/₹$!@%*\s]*$/;


  const handleValidatedInput = (
    value: string,
    updateFn: (val: string) => void,
  ) => {
    if (!allowedRegex.test(value)) return;
    updateFn(value);
  };

  const handleValidatedPaste = (
    e: React.ClipboardEvent<HTMLInputElement | HTMLTextAreaElement>,
  ) => {
    const paste = e.clipboardData.getData("text");
    if (!allowedRegex.test(paste)) {
      e.preventDefault();
    }
  };

  return (
    <section className={`${styles.editNews} py-4`}>
      <Card className={`${styles.cardCommon} p-1 p-sm-3 p-md-4`}>
        <Card.Body>
          <h4 className={styles.pageTitle}>{"Edit Broadcast"}</h4>

          <Form>
            <Row className="g-3">
              {/* TITLE */}
              <Col xs={12}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Title <span className={styles.required}>*</span>
                  </Form.Label>
                  <Form.Control
                    className={styles.inputBox}
                    value={form.Title || ""}
                    maxLength={255}
                    disabled={!canEditFields}
                    onChange={(e) => {
                      if (!canEditFields) return;

                      handleValidatedInput(e.target.value, (val) =>
                        setForm({ ...form, Title: val }),
                      );
                    }}
                    onPaste={handleValidatedPaste}
                  />
                  <small style={{ color: (form.Title || "").length > 250 ? "red" : "#666" }}>
  {(form.Title || "").length}/255 characters
</small>
                </Form.Group>
              </Col>

              {/* Dl group */}
              <Col md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    DL Group <span className={styles.required}>*</span>
                  </Form.Label>

                  {!canEditFields ? (
                    <Form.Control
                      className={styles.inputBox}
                      disabled
                      value={
                        dlGroups && dlGroups.length > 0
                          ? dlGroups.map((g) => g.label).join(", ")
                          : ""
                      }
                    />
                  ) : (
                    <div>
                      <PeoplePickerCustom
                        context={context}
                        isMulti
                        selectedValue={dlGroups}
                        // onChange={(val: Option[] | null) => {
                        //   if (!canEditFields) return;

                        //   // UI update
                        //   setDlGroups(val || []);

                        //   // Save to SharePoint
                        //   setForm({
                        //     ...form,
                        //     DLGroupId: {
                        //       results: val?.map((v) => Number(v.value)) || [],
                        //     },
                        //   });
                        // }}
                        onChange={(val: Option | Option[] | null) => {
  if (!canEditFields) return;

  const normalized = !val
    ? []
    : Array.isArray(val)
    ? val
    : [val];

  setDlGroups(normalized);

  setForm({
    ...form,
    DLGroupId: {
      results: normalized.map((v) => Number(v.value)),
    },
  });
}}  
                      />
                    </div>
                  )}
                </Form.Group>
              </Col>

              {/* SEND MAIL */}
              <Col md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Send Mail <span className={styles.required}>*</span>
                  </Form.Label>
                  <Form.Select
                    className={styles.inputBox}
                    value={form.Sendmail ? "Yes" : "No"}
                    disabled={!canEditFields}
                    onChange={(e) => {
                      if (!canEditFields) return;
                      setForm({ ...form, Sendmail: e.target.value === "Yes" });
                    }}
                  >
                    <option>No</option>
                    <option>Yes</option>
                  </Form.Select>
                </Form.Group>
              </Col>

              {/* EVENT TYPE */}

              <Col xs={12} sm={6} md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    BroadCast Type <span className={styles.required}>*</span>
                  </Form.Label>

                  {canEditFields ? (
                    <TaxonomyPicker
                      label=""
                      context={context}
                      allowMultipleSelections={false}
                      // termsetNameOrID={termSetBroadcastTypesId}
                      termsetNameOrID="7b5de19a-32f9-4e96-bcc4-ede23914510b"
                      anchorId="d6c3459f-1e92-47dd-be09-9c9b747e61e1"
                      panelTitle="Select Broadcast Type"
                      initialValues={broadCastType ? [broadCastType] : []}
                      onChange={(terms) => {
                        if (!terms || terms.length === 0) {
                          setBroadCastType(null);
                          setForm({ ...form, BroadcastType: null });
                          return;
                        }

                        const t = terms[0];

                        setBroadCastType(t);

                        setForm({
                          ...form,
                          BroadcastType: {
                            Label: t.name,
                            TermGuid: t.key,
                          },
                        });
                      }}
                      isTermSetSelectable={false}
                    />
                  ) : (
                    <Form.Control
                      className={styles.inputBox}
                      disabled
                      value={form.BroadcastType?.Label || ""}
                    />
                  )}
                </Form.Group>
              </Col>

              {/* RELATED URL */}
              <Col md={8}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Related Site URL
                    {/* removed * since not mandatory */}
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    value={form.RelatedSiteURL || ""}
                    disabled={!canEditFields}
                    onChange={(e) => {
                      if (!canEditFields) return;

                      const value = e.target.value.trim();
                      setForm({ ...form, RelatedSiteURL: value });

                      // Validate ONLY if user entered something
                      if (value && !isValidUrl(value)) {
                        setUrlError("Enter valid URL (http/https only).");
                      } else {
                        setUrlError("");
                      }
                    }}
                  />

                  {urlError && (
                    <div
                      style={{
                        color: "red",
                        fontSize: "12px",
                        marginTop: "2px",
                      }}
                    >
                      {urlError}
                    </div>
                  )}
                </Form.Group>
              </Col>

              {/* ConfidentialityIndex */}
              <Col xs={12} sm={12} md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Confidentiality Index{" "}
                    {/* <span className={styles.required}>*</span> */}
                  </Form.Label>

                  <Form.Control
                    disabled
                    className={styles.inputBox}
                    type="text"
                    value={form.ConfidentialityIndex?.Label || "Confidential"}
                    readOnly
                  />
                </Form.Group>
              </Col>

              {/* ArchivalPolicy */}
              <Col xs={12} sm={6} md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Archival Policy
                    {/* <span className={styles.required}>*</span> */}
                  </Form.Label>

                  <Form.Control
                    disabled
                    className={styles.inputBox}
                    type="text"
                    value={
                      form.ArchivalPolicy?.Label || "As Per Company Policy"
                    }
                    readOnly
                  />
                </Form.Group>
              </Col>

              {/* EVENT FROM */}

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="eventDate" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Event Date<span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="date"
                    value={form.EventDate || ""}
                    disabled={!canEditFields}
                    onChange={(e) => {
                      if (!canEditFields) return;
                      setForm({ ...form, EventDate: e.target.value });
                    }}
                    required
                  />
                </Form.Group>
              </Col>

              {/* EVENT TO */}

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="expiryDate" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Expiry Date
                    {/* <span className={styles.required}>*</span> */}
                  </Form.Label>

                  <Form.Control
                    disabled
                    className={styles.inputBox}
                    type="date"
                    value={form.ExpiryDate || getExpiryPreview()}
                    readOnly
                  />
                </Form.Group>
              </Col>

              {/* APPROVER */}
              <Col md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Approver <span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Select
                    className={styles.inputBox}
                    value={form.AssignTo?.[0] || ""}
                    disabled={disableWorkflowDropdowns}
                    onChange={(e) => {
                      const val = e.target.value;

                      setForm({
                        ...form,
                        AssignTo: val ? [Number(val)] : [],
                      });
                    }}
                  >
                    <option value="">Select</option>
                    {approvers.map((u) => (
                      <option key={u.Id} value={u.Id}>
                        {u.Title}
                      </option>
                    ))}
                  </Form.Select>
                </Form.Group>
              </Col>

              {/* CONTENT MANAGER */}
              <Col md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Content Manager <span className={styles.required}>*</span>
                  </Form.Label>
                  <Form.Select
                    className={styles.inputBox}
                    value={selectedContentManager}
                    disabled={disableWorkflowDropdowns}
                    onChange={(e) => {
                      const val = e.target.value;
                      setSelectedContentManager(val);

                      setForm({
                        ...form,
                        FollowOnActionBy: val ? [Number(val)] : [],
                      });
                    }}
                  >
                    <option value="">Select</option>
                    {contentManagers.map((u) => (
                      <option key={u.Id} value={u.Id}>
                        {u.Title}
                      </option>
                    ))}
                  </Form.Select>
                </Form.Group>
              </Col>

              {/* VIEW BY */}
              <Col md={4}>
                <Form.Label className={styles.floatingLabelStyle}>
                  View By
                </Form.Label>
                {canEditFields ? (
                  <PeoplePickerCustom
                    context={context}
                    isClearable={true}
                    selectedValue={viewBy}
                    onChange={(val) =>
                      setForm({
                        ...form,
                        ViewBy: val ? [Number((val as any).value)] : [],
                      })
                    }
                  />
                ) : (
                  <Form.Group>
                    <Form.Control
                      className={styles.inputBox}
                      disabled
                      value={viewBy?.label || ""}
                    />
                  </Form.Group>
                )}
              </Col>

              {/*  Upload Related Documents */}
              <Row className="g-3">
                <Col xs={12} md={12}>
                  <Form.Group controlId="relatedDocuments">
                    <Form.Label className={styles.floatingLabelStyle}>
                      Upload Related Documents{" "}
                      <span className={styles.required}>*</span>
                      <small>
                        ( Please select at least 1 file. The first file will
                        become the main file. Document size must be less than 4
                        MB (pdf, doc, docx). Image size must be less than 2 MB
                        (bmp, gif, png, jpg, jpeg ))
                      </small>
                    </Form.Label>

                    <Form.Control
                      className={styles.inputBox}
                      type="file"
                      multiple
                      accept=".pdf,.doc,.docx,.bmp,.png,.jpg,.jpeg,.gif"
                      onChange={handleFileChange}
                      disabled={!canEditFields}
                    />
                    {/* Uploaded file list */}
                    {/* Pending Upload (Local files before save) */}
                    {/* ================= ATTACHMENTS ================= */}
                    {(uploadedFiles.length > 0 ||
                      existingAttachments.length > 0) && (
                      <div className="mt-3">
                        {/*  Pending (not yet saved) */}
                        {uploadedFiles.length > 0 && (
                          <>
                            <strong>Selected Files</strong>
                            {uploadedFiles.map((file, index) => (
                              <div
                                key={`${file.name}-${index}`}
                                className="d-flex justify-content-between align-items-center border rounded px-2 py-1 mb-1"
                              >
                                <span className="text-truncate">
                                  {file.name}
                                </span>

                                {/* {!isReadOnly && ( */}
                                {canEditFields && (
                                  <Button
                                    variant="link"
                                    className="text-danger p-0"
                                    onClick={() =>
                                      setUploadedFiles((prev) =>
                                        prev.filter((_, i) => i !== index),
                                      )
                                    }
                                  >
                                    Delete
                                  </Button>
                                )}
                              </div>
                            ))}
                          </>
                        )}

                        {/* Saved attachments */}
                        {existingAttachments.length > 0 && (
                          <>
                            <strong className="mt-3 d-block">
                              Uploaded Files
                            </strong>
                            {existingAttachments.map((file) => (
                              <div
                                key={file.FileName}
                                className="d-flex justify-content-between align-items-center border rounded px-2 py-1 mb-1"
                              >
                                <a
                                  href={`${context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(file.ServerRelativeUrl)}`}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                >
                                  {file.FileName}
                                </a>

                                {canEditFields && (
                                  <Button
                                    variant="link"
                                    className="text-danger p-0"
                                    onClick={() =>
                                      deleteExistingAttachment(file.FileName)
                                    }
                                  >
                                    Delete
                                  </Button>
                                )}
                              </div>
                            ))}
                          </>
                        )}
                      </div>
                    )}
                    {/* Help / Validation Text BELOW input */}
                    <div
                      style={{
                        fontSize: "12px",
                        color: "#6c757d",
                        marginTop: "4px",
                      }}
                    >
                      If the first attachement is image then it will go as a
                      email body in broadcast mail and all other will go as
                      attachement.
                    </div>
                  </Form.Group>
                </Col>
              </Row>

              {/* ThumbnailCaption */}
              <Row className="mt-1 g-3">
                <Col xs={12} md={12}>
                  <Form.Group controlId="thumbnailCaption">
                    <Form.Label className={styles.floatingLabelStyle}>
                      Image Description{" "}
                      <span className={styles.required}>*</span>
                    </Form.Label>

                    <Form.Control
                      as="textarea"
                      ref={commentRef}
                      rows={2}
                      className={styles.inputBox}
                      placeholder="Image Description"
                      value={form.ThumbnailCaption || ""}
                      disabled={!canEditFields}
                      style={{ resize: "none", overflow: "hidden" }}
                      onChange={(e) => {
                        if (!canEditFields) return;

                        e.target.style.height = "auto";
                        e.target.style.height = e.target.scrollHeight + "px";

                        handleValidatedInput(e.target.value, (val) =>
                          setForm({ ...form, ThumbnailCaption: val }),
                        );
                      }}
                      onPaste={handleValidatedPaste}
                    />
                    <div
                      style={{
                        fontSize: "12px",
                        color: "#6c757d",
                        marginTop: "4px",
                      }}
                    >
                      (Images cannot be read by visually impaired people. Please
                      describe/write the text for images in this text box).
                    </div>
                  </Form.Group>
                </Col>
              </Row>

              {/* Additional info + Comments */}

              <Row className="g-3">
                <Col xs={12} sm={6} md={4}>
                  <Form.Group controlId="additionalInfo1" className="">
                    <Form.Label className={styles.floatingLabelStyle}>
                      Additional Info 1
                    </Form.Label>

                    <Form.Control
                      className={styles.inputBox}
                      type="text"
                      value={form.AdditionalInformation1 || ""}
                      disabled={!canEditFields}
                      onChange={(e) => {
                        if (!canEditFields) return;

                        handleValidatedInput(e.target.value, (val) =>
                          setForm({
                            ...form,
                            AdditionalInformation1: val,
                          }),
                        );
                      }}
                      onPaste={handleValidatedPaste}
                    />
                  </Form.Group>
                </Col>

                <Col xs={12} sm={6} md={4}>
                  <Form.Group controlId="additionalInfo2" className="">
                    <Form.Label className={styles.floatingLabelStyle}>
                      Additional Info 2
                    </Form.Label>

                    <Form.Control
                      className={styles.inputBox}
                      type="text"
                      value={form.AdditionalInformation2 || ""}
                      disabled={!canEditFields}
                      onChange={(e) => {
                        if (!canEditFields) return;

                        handleValidatedInput(e.target.value, (val) =>
                          setForm({
                            ...form,
                            AdditionalInformation2: val,
                          }),
                        );
                      }}
                      onPaste={handleValidatedPaste}
                    />
                  </Form.Group>
                </Col>

                <Col xs={12} sm={6} md={4}>
                  <Form.Group controlId="additionalInfo3" className="">
                    <Form.Label className={styles.floatingLabelStyle}>
                      Additional Info 3
                    </Form.Label>

                    <Form.Control
                      className={styles.inputBox}
                      type="text"
                      value={form.AdditionalInformation3 || ""}
                      disabled={!canEditFields}
                      onChange={(e) => {
                        if (!canEditFields) return;

                        handleValidatedInput(e.target.value, (val) =>
                          setForm({
                            ...form,
                            AdditionalInformation3: val,
                          }),
                        );
                      }}
                      onPaste={handleValidatedPaste}
                    />
                  </Form.Group>
                </Col>
              </Row>

              <Row className="g-3">
                <Col>
                  <Form.Group controlId="comments" className="">
                    <Form.Label className={styles.floatingLabelStyle}>
                      Comment
                    </Form.Label>

                    <Form.Control
                      as="textarea"
                      ref={commentRef}
                      value={form.Comments || ""}
                      disabled={!canEditComment}
                      style={{
                        minHeight: "50px",
                        resize: "none",
                        overflow: "hidden",
                      }}
                      onChange={(e) => {
                        if (!canEditComment) return;

                        e.target.style.height = "auto";
                        e.target.style.height = `${Math.max(e.target.scrollHeight, 50)}px`;

                        handleValidatedInput(e.target.value, (val) =>
                          setForm({ ...form, Comments: val }),
                        );
                      }}
                      onPaste={handleValidatedPaste}
                    />
                  </Form.Group>
                </Col>
              </Row>
            </Row>

            {/* ================= BUTTONS ================= */}

            {/* Buttons */}
            <Row className="mt-4">
              <div className="d-flex flex-column flex-md-row gap-2">
                {/* Preview */}
                <Button
                  className={`${styles.primaryOutlineBtn} btn`}
                  onClick={onPreviewPopup}
                >
                  Preview
                </Button>

                {/* Save */}
                {canShowSaveButton && (
                  <Button
                    className={`${styles.primaryBtn} btn`}
                    // onClick={onSaveDraft}
                    onClick={async () => {
                      if (urlError) {
                        setPopupType("danger");
                        setPopupMessage(urlError);
                        setShowPopup(true);
                        return;
                      }
                      await onSaveDraft();
                    }}
                  >
                    Save
                  </Button>
                )}

                <div className="ms-md-auto d-flex flex-column flex-md-row gap-2 px-md-3">
                  {/* Approve / Reject / Reassign */}
                  {canShowApproveRejectButtons && (
                    <>
                      <Button
                        className={styles.reassignBtn}
                        onClick={onReAssign}
                      >
                        ReAssign
                      </Button>

                      <Button className={styles.approveBtn} onClick={onApprove}>
                        Approve
                      </Button>

                      <Button className={styles.rejectBtn} onClick={onReject}>
                        Reject
                      </Button>
                    </>
                  )}

                  {/* Delete */}
                  {canShowDeleteButton && (
                    <Button
                      variant="outline-danger"
                      className={styles.deleteBtn}
                      onClick={onDelete}
                    >
                      Delete
                    </Button>
                  )}

                  {/* Cancel */}
                  <Button
                    className={`${styles.primaryOutlineBtn} btn`}
                    onClick={onBack}
                  >
                    Cancel
                  </Button>

                  {/* Submit */}
                  {canShowSubmitButton && (
                    <Button
                      className={`${styles.primaryBtn} btn`}
                      // onClick={onSubmit}
                      onClick={async () => {
                        if (urlError) {
                          setPopupType("danger");
                          setPopupMessage(urlError);
                          setShowPopup(true);
                          return;
                        }
                        await onSubmit();
                      }}
                    >
                      Submit
                    </Button>
                  )}
                </div>
              </div>
            </Row>
          </Form>
        </Card.Body>
      </Card>

      {/* Popup Modal */}
      <PopupModal
        show={showPopup}
        type={popupType}
        message={popupMessage}
        onClose={async (result?: "yes" | "no") => {
          // Cancel clicked
          if (popupType === "confirm" && result === "no") {
            setConfirmAction(null);
            setShowPopup(false);
            return;
          }

          // Confirm YES clicked
          if (popupType === "confirm" && result === "yes" && confirmAction) {
            const action = confirmAction;
            setConfirmAction(null); // 🔹 clear first
            setShowPopup(false); // 🔹 close confirm popup FIRST
            await action(); // 🔹 run reassignment → will show SUCCESS popup
            return;
          }

          // Normal close (OK button)
          setShowPopup(false);

          if (popupType === "success") {
            if (refreshList) await refreshList();
            onBack();
          }
        }}
      />
    </section>
  );
};

export default BroadcastEditForm;
