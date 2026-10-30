import * as React from "react";
import { useState, useEffect } from "react";
//import "@fontsource/inter";
import "@fontsource/inter/500.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/700.css";
import { TaxonomyPicker } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import styles from "./MyViewBroadcast.module.scss";
import { Form, Row, Col, Card, Button } from "react-bootstrap";

import PeoplePicker, { Option } from "../services/PeoplePicker";
import PopupModal from "../services/PopupModel";
import { BroadcastService } from "../services/BroadcastService";
import ErrorService from "../services/ErrorService";


interface Props {
  context: any;
  service: BroadcastService;
  termSetBroadcastTypesId: string;
  onBack: () => void;
  onEditAfterSave: (id: number) => void;
}

const CreateBroadcast: React.FC<Props> = ({
  context,
  service,
  termSetBroadcastTypesId,
  onBack,
  onEditAfterSave,
}) => {
  /* ================= STATE ================= */

  const [title, setTitle] = useState("");

  const [dlGroups, setDlGroups] = useState<Option[]>([]);

  const [sendMail, setSendMail] = useState(false);

  const [relatedUrl, setRelatedUrl] = useState("");
  // const [confidentialIndex, setConfidentialIndex] = useState("");
  // const [archivalPolicy, setArchivalPolicy] = useState("");

  const [approvers, setApprovers] = useState<any[]>([]);
  const [selectedApprover, setSelectedApprover] = useState<string>("");
  const [contentManagers, setContentManagers] = useState<any[]>([]);
  const [selectedContentManager, setSelectedContentManager] =
    useState<string>("");

  const [eventFrom, setEventFrom] = useState<string>("");
  const [addInfo1, setAddInfo1] = useState<string>("");
  const [addInfo2, setAddInfo2] = useState<string>("");
  const [addInfo3, setAddInfo3] = useState<string>("");
  const [comments, setComments] = useState<string>("");
  //const [thumbnailFile, setThumbnailFile] = useState<File | null>(null);
  const [thumbnailCaption, setThumbnailCaption] = useState<string>("");

  // ===== Taxonomy =====
  const [broadCastType, setBroadCastType] = useState<any>(null);

  // ===== Attachment Save Id =====
  const [itemId, setItemId] = useState<number | null>(null);

  const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
  const [existingAttachments, setExistingAttachments] = useState<any[]>([]);

  const [viewBy, setViewBy] = useState<Option | null>(null);

  const [department, setDepartment] = useState<string>("");
  const [urlError, setUrlError] = useState<string>("");

  const [showPopup, setShowPopup] = useState(false);
  const [popupMessage, setPopupMessage] = useState("");
  const [popupType, setPopupType] = useState<"success" | "danger">("success");

  const [savedId, setSavedId] = useState<number | null>(null);
      const [isSaving, setIsSaving] = useState(false);

  /* ================= LOAD METADATA ================= */

    const confidentialTerm = {
  Label: "Confidential",
  TermGuid: "78e3abf4-3b6a-4fab-ba1c-baca497498cd",
  WssId: -1,
};

const archivalTerm = {
  Label: "As Per Company Policy",
  TermGuid: "93858a26-566e-485e-b8ac-8a480ffcf903",
  WssId: -1,
};

  useEffect(() => {
  const loadData = async () => {
    try {
      //  Get profile (for department display only)
      const profileInfo = await service.getLoggedInUserProfileInfo();
      setDepartment(profileInfo?.department || "");

      //  Validate user (SBU OR Department handled inside)
      const isValid = await service.isUserPartOfSbuTeam();

      if (!isValid) {
        setApprovers([]);
        setContentManagers([]);
        return;
      }

      //  Get approvers (auto handles SBU + Dept + priority)
      let appr = await service.getApproversByTeam();

      //  Add Reporting Manager (if not exists)
      const managerUser = profileInfo?.reportingManagerUser;

      if (managerUser?.Id) {
        const exists = appr.some((u: any) => u.Id === managerUser.Id);
        if (!exists) {
          appr.unshift(managerUser);
        }
      }

      //  Remove current user
      const currentUserId = await service.getCurrentUserId();

      const filteredApprovers = appr.filter(
        (u: any) => u.Id !== currentUserId
      );

      setApprovers(filteredApprovers);

      //  Get Content Managers
      const cms = await service.getContentManagersByTeam();
      setContentManagers(cms);

    } catch (err) {

  ErrorService.logError(
    "CreateBroadcast.loadData",
    err
  );    setPopupType("danger");

  setPopupMessage(
    ErrorService.getUserMessage(err)
  );

  setShowPopup(true); 
 }
  };

  loadData();
}, []);

  /* ================= SAVE ================= */
  const dlGroupPayload = async (): Promise<number[]> => {
    return await Promise.all(
      dlGroups.map(async (g) => {
        //  Use existing ID directly (EDIT case)
        if (typeof g.value === "number") {
          return g.value;
        }

        //  Only ensure when loginName is a real claim
        if (g.loginName && g.loginName.includes("|")) {
          return await service.ensurePrincipal(g.loginName);
        }

        throw new Error(
          `Invalid DL Group. Cannot resolve: ${g.label || g.loginName}`,
        );
      }),
    );
  };

  const isValidUrl = (url: string): boolean => {
    if (!url) return false;

    const pattern = new RegExp(
      "^(https?:\\/\\/)" + // require http/https
        '((([a-zA-Z0-9$-_@.&+!*"(),]|(%[0-9a-fA-F][0-9a-fA-F]))+\\.)+' +
        "([a-zA-Z]{2,}))" + // domain
        "(\\:[0-9]{1,5})?" + // port
        "(\\/.*)?$",
      "i",
    );

    return pattern.test(url.trim());
  };

  // validation msg
  const showError = (msg: string) => {
    window.scrollTo({ top: 0, behavior: "smooth" });
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
    if (!itemId) return;
    await service.deleteAttachment(itemId, fileName);
    setExistingAttachments(await service.getAttachments(itemId));
  };

  const handleSave = async () => {
    /* ---------- BASIC REQUIRED ---------- */
          if (isSaving) return;


    if (!title?.trim()) return showError("Title is required");

     if (title.trim().length > 255) {
    return showError("Title should not exceed 255 characters.");
  } 

     setIsSaving(true);

    try {
      const dlGroup = await dlGroupPayload();

      const result = await service.saveBroadcast(
        title,
        dlGroup,
        sendMail,
        relatedUrl,
        //confidentialIndex,
        eventFrom,
        //archivalPolicy,
        // eventTo,
        addInfo1,
        addInfo2,
        addInfo3,
        comments,
        thumbnailCaption,
        selectedApprover,
        selectedContentManager,
        broadCastType
          ? { name: broadCastType.name, key: broadCastType.key }
          : null,
        viewBy?.value || null,
        department,
                confidentialTerm,
  archivalTerm 

      ); 

      //  Save itemId for attachments
      setItemId(result.itemId);

      //  Upload attachments AFTER item created
      if (uploadedFiles.length > 0) {
        await service.uploadAttachments(result.itemId, uploadedFiles);
        setUploadedFiles([]);
      }

      //  Reload attachments
      setExistingAttachments(await service.getAttachments(result.itemId));

      setPopupType("success");
      setPopupMessage(`Broadcast ${result.requestNumber} saved successfully.`);
      setSavedId(result.itemId);
      setShowPopup(true);

      /* Reset */
      setTitle("");
      setDlGroups([]);
      setSendMail(false);
      setRelatedUrl("");
      setEventFrom("");
      setSelectedApprover("");
      setSelectedContentManager("");
    } catch (e) {
 ErrorService.logError("CreateBroadcast.handleSave", e);
      setPopupType("danger");
setPopupMessage(
  ErrorService.getUserMessage(e)
);       setShowPopup(true);
    }

    
finally {
    setIsSaving(false);
  }

  };

  const getExpiryPreview = () => {
    const dt = new Date();
    dt.setDate(dt.getDate() + 90);

    const yyyy = dt.getFullYear();
    const mm = ("0" + (dt.getMonth() + 1)).slice(-2);
    const dd = ("0" + dt.getDate()).slice(-2);

    return `${yyyy}-${mm}-${dd}`;
  };

 
  /* ================= VALIDATION HELPER ================= */

  // Allowed characters (A–Z a–z 0–9 + approved special chars)
  //const allowedRegex = /^[A-Za-z0-9()&_\-+=\[\]{}:;'",.?/₹$!@%*\s]*$/;
  const allowedRegex = /^[A-Za-z0-9\u0900-\u097F()&_\-+=\[\]{}:;'",.?/₹$!@%*\s]*$/;


  // Reusable change handler
  const handleValidatedInput = (
    value: string,
    updateFn: (val: string) => void,
  ) => {
    if (!allowedRegex.test(value)) return;
    updateFn(value);
  };

  // Paste protection
  const handleValidatedPaste = (
    e: React.ClipboardEvent<HTMLInputElement | HTMLTextAreaElement>,
  ) => {
    const paste = e.clipboardData.getData("text");
    if (!allowedRegex.test(paste)) {
      e.preventDefault();
    }
  };

  /* ================= UI ================= */

  return (
    <section className={`${styles.createNews} py-4`}>
      <Card className={`${styles.cardCommon} p-1 p-sm-3 p-md-4`}>
        <Card.Body>
          <h4 className={styles.pageTitle}>Create BroadCast</h4>

          <Form>
            {/* TITLE */}
            <Row className="g-3">
              <Col xs={12}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Title <span className={styles.required}>*</span>
                  </Form.Label>
                  <Form.Control
                    className={styles.inputBox}
                    value={title}
                    maxLength={255}
                    // onChange={(e) => setTitle(e.target.value)}
                    onChange={(e) =>
                      handleValidatedInput(e.target.value, setTitle)
                    }
                    onPaste={handleValidatedPaste}
                  />
                   <small style={{ color: title.length > 250 ? "red" : "#666" }}>
    {title.length}/255 characters
  </small>
                </Form.Group>
              </Col>

              {/* DL GROUP */}
              <Col md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    DL Group <span className={styles.required}>*</span>
                  </Form.Label>

                  <div>
                    <PeoplePicker
                      context={context}
                      isMulti={true}
                      selectedValue={dlGroups}
                      // onChange={(val: Option[] | null) =>
                      //   setDlGroups(val || [])
                      // }
                       onChange={(val: Option | Option[] | null) => {
  if (Array.isArray(val)) {
    setDlGroups(val);
  } else if (val) {
    setDlGroups([val]); // convert single to array
  } else {
    setDlGroups([]);
  }
}} 
                    />
                  </div>
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
                    value={sendMail ? "Yes" : "No"}
                    onChange={(e) => setSendMail(e.target.value === "Yes")}
                  >
                    <option value="No">No</option>
                    <option value="Yes">Yes</option>
                  </Form.Select>
                </Form.Group>
              </Col>

              {/* EVENT TYPE */}

              <Col xs={12} sm={6} md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    BroadCast Type <span className={styles.required}>*</span>
                  </Form.Label>
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
                      if (terms && terms.length > 0) {
                        setBroadCastType(terms[0]);
                      } else {
                        setBroadCastType(null);
                      }
                    }}
                    isTermSetSelectable={false}
                  />
                </Form.Group>
              </Col>

              {/* RELATED URL */}
              <Col md={8}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Related Site URL
                    {/* remove star since not mandatory */}
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    value={relatedUrl}
                    onChange={(e) => {
                      const value = e.target.value.trim();
                      setRelatedUrl(value);

                      // Only validate if user entered something
                      if (value && !isValidUrl(value)) {
                        setUrlError(
                          "Please enter a valid URL (http/https only).",
                        );
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

              {/* CONFIDENTIAL */}
              <Col md={4}>
                <Form.Group controlId="confidentialityIndex" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Confidentiality Index{" "}
                    {/* <span className={styles.required}>*</span> */}
                  </Form.Label>
                  <Form.Control
                    disabled
                    className={styles.inputBox}
                    type="text"
                    value={"Confidential"}
                    readOnly
                  />
                </Form.Group>
              </Col>

              {/* ARCHIVAL */}
              <Col md={4}>
                <Form.Group controlId="archivalPolicy" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Archival Policy
                    {/* <span className={styles.required}>*</span> */}
                  </Form.Label>
                  <Form.Control
                    disabled
                    className={styles.inputBox}
                    type="text"
                    value={"As Per Company Policy"}
                    readOnly
                  />
                </Form.Group>
              </Col>

              {/* FROM DATE */}
              <Col md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Event Date <span className={styles.required}>*</span>
                  </Form.Label>
                  <Form.Control
                    type="date"
                    className={styles.inputBox}
                    value={eventFrom}
                    onChange={(e) => setEventFrom(e.target.value)}
                  />
                </Form.Group>
              </Col>

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
                    value={getExpiryPreview()}
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
                    required
                    value={selectedApprover}
                    onChange={(e) => setSelectedApprover(e.target.value)}
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
                    required
                    value={selectedContentManager}
                    onChange={(e) => setSelectedContentManager(e.target.value)}
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

              {/* View By */}
              <Col xs={12} sm={6} md={4}>
                <Form.Label className={styles.floatingLabelStyle}>
                  View By
                </Form.Label>
                <PeoplePicker
                  context={context}
                  isClearable={true}
                  selectedValue={viewBy}
                  onChange={(val) =>
                    setViewBy(Array.isArray(val) ? (val[0] ?? null) : val)
                  }
                />
              </Col>
            </Row>
            {/* Thumbnail */}
            <Row className="mt-1 g-3">
              <Col xs={12} md={12}>
                <Form.Group controlId="relatedDocuments">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Upload Related Documents{" "}
                    <span className={styles.required}>*</span>
                    <small>
                      ( Please select at least 1 file. The first file will
                      become the main file. Document size must be less than 4 MB
                      (pdf, doc, docx). Image size must be less than 2 MB (bmp,
                      gif, png, jpg, jpeg ))
                    </small>
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="file"
                    multiple
                    accept=".pdf,.doc,.docx,.bmp,.png,.jpg,.jpeg,.gif"
                    onChange={handleFileChange}
                    //disabled={isReadOnly}
                  />

                  {/* Uploaded file list */}
                  {/* Pending Upload (Local files before save) */}
                  {/* ================= ATTACHMENTS ================= */}
                  {(uploadedFiles.length > 0 ||
                    existingAttachments.length > 0) && (
                    <div className="mt-3">
                      {/* 🔹 Pending (not yet saved) */}
                      {uploadedFiles.length > 0 && (
                        <>
                          <strong>Selected Files</strong>
                          {uploadedFiles.map((file, index) => (
                            <div
                              key={`${file.name}-${index}`}
                              className="d-flex justify-content-between align-items-center border rounded px-2 py-1 mb-1"
                            >
                              <span className="text-truncate">{file.name}</span>

                              {/* {!isReadOnly && ( */}
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
                              {/* )} */}
                            </div>
                          ))}
                        </>
                      )}

                      {/* 🔹 Saved attachments */}
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
                                href={file.ServerRelativeUrl}
                                target="_blank"
                                rel="noreferrer"
                                className="text-truncate"
                              >
                                {file.FileName}
                              </a>

                              {/* {!isReadOnly && ( */}
                              <Button
                                variant="link"
                                className="text-danger p-0"
                                onClick={() =>
                                  deleteExistingAttachment(file.FileName)
                                }
                              >
                                Delete
                              </Button>
                              {/* )} */}
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
                    If the first attachement is image then it will go as a email
                    body in broadcast mail and all other will go as attachment.
                  </div>
                </Form.Group>
              </Col>
            </Row>

            <Row className="mt-1 g-3">
              <Col xs={12} md={12}>
                <Form.Group controlId="thumbnailCaption">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Image Description <span className={styles.required}>*</span>
                  </Form.Label>
                  <Form.Control
                    as="textarea"
                    rows={2}
                    className={styles.inputBox}
                    placeholder="Image Description"
                    style={{ resize: "none", overflow: "hidden" }}
                    value={thumbnailCaption}
                    onChange={(e) => {
                      e.target.style.height = "auto";
                      e.target.style.height = e.target.scrollHeight + "px";

                      handleValidatedInput(e.target.value, setThumbnailCaption);
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

            {/* Additional Info & Comment */}

            <Row className="mt-1 g-3">
              {["1", "2", "3"].map((num, index) => (
                <Col key={num} xs={12} sm={6} md={4}>
                  <Form.Group controlId={`additionalInfo${num}`} className="">
                    <Form.Label className={styles.floatingLabelStyle}>
                      {`Additional Info ${num}`}
                    </Form.Label>

                    <Form.Control
                      className={styles.inputBox}
                      type="text"
                      placeholder={`Additional Info ${num}`}
                      value={
                        index === 0
                          ? addInfo1
                          : index === 1
                            ? addInfo2
                            : addInfo3
                      }
                      onChange={(e) => {
                        const value = e.target.value;

                        handleValidatedInput(value, (val) => {
                          if (index === 0) setAddInfo1(val);
                          if (index === 1) setAddInfo2(val);
                          if (index === 2) setAddInfo3(val);
                        });
                      }}
                      onPaste={handleValidatedPaste}
                    />
                  </Form.Group>
                </Col>
              ))}
            </Row>

            <Row className="mt-1 g-3">
              <Col>
                <Form.Group controlId="comment" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Comment
                  </Form.Label>

                  <Form.Control
                    as="textarea"
                    placeholder="Comment"
                    value={comments}
                    style={{
                      minHeight: "50px",
                      resize: "none",
                      overflow: "hidden",
                    }}
                    onChange={(e) => {
                      e.target.style.height = "auto";
                      e.target.style.height = `${Math.max(e.target.scrollHeight, 50)}px`;

                      handleValidatedInput(e.target.value, setComments);
                    }}
                    onPaste={handleValidatedPaste}
                  />
                </Form.Group>
              </Col>
            </Row>

            {/* BUTTONS */}
            <Row className="mt-4">
              <div className="d-flex justify-content-end gap-2">
                <Button
                  className={`${styles.primaryOutlineBtn}`}
                  onClick={onBack}
                >
                  Cancel
                </Button>

                <Button className={`${styles.primaryBtn}`} onClick={handleSave} disabled={isSaving}
>
  {isSaving ? "Saving..." : "Save"}
</Button>
              </div>
            </Row>
          </Form>
        </Card.Body>
      </Card>

      <PopupModal
        show={showPopup}
        type={popupType}
        message={popupMessage}
        onClose={() => {
          setShowPopup(false);
          if (popupType === "success" && savedId) {
            onEditAfterSave(savedId);
          }
        }}
      />
    </section>
  );
};

export default CreateBroadcast;
