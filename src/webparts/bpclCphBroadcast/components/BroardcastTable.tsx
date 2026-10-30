import * as React from "react";
import { useMemo, useState, useEffect } from "react";

import {
  ButtonGroup,
  Card,
  Form,
  Pagination,
  Table,
  ToggleButton,
  Button,
} from "react-bootstrap";

import styles from "./MyViewBroadcast.module.scss";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap-icons/font/bootstrap-icons.css";

import Container from "react-bootstrap/Container";
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";
import Tabs from "react-bootstrap/Tabs";
import Tab from "react-bootstrap/Tab";

import newsIcon from "../assets/news.png";
import eventsIcon from "../assets/events.png";
import broadcastIcon from "../assets/broadcasts.png";
import brandIcon from "../assets/brands.png";

import { BroadcastService } from "../services/BroadcastService";
import ErrorService from "../services/ErrorService";


const STATUSES = [
  "Draft",
  "Pending",
  "Published",
  "Approval Queue",
  "View Requests",
];

interface Props {
  context: any;
  items: any[];
    isLoading: boolean;

  selectedStatus: string;
  search: string;
  onStatusChange: (v: string) => void;
  onSearchChange: (v: string) => void;
  onEdit: (item: any) => void;
  onCreate: () => void;
  canCreateBroadcast: boolean;
  onPreviewPopup: (item: any) => void;
  isSiteAdmin?: boolean;
}

const PAGE_SIZES = [10, 25, 50, 100];

const BroadcastTable: React.FC<Props> = ({
  context, // ADD
  items,
      isLoading, 

  selectedStatus,
  search,
  onStatusChange,
  onSearchChange,
  onEdit,
  onCreate,
  canCreateBroadcast,
  onPreviewPopup,
  isSiteAdmin,
}) => {
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState<number>(10);

  const [activeMainTab, setActiveMainTab] = useState("broadcasts");

  const [currentUserId, setCurrentUserId] = useState<number>(0);
    const [isUserLoading, setIsUserLoading] = useState(true);

      const broadcastService = React.useMemo(
    () => new BroadcastService(context),
    [context],
  );

  // useEffect(() => {
  //   (async () => {
  //     const user = await broadcastService.sp.web.currentUser();
  //     setCurrentUserId(user.Id);
  //   })();
  // }, []);

      useEffect(() => {
  const loadCurrentUser = async () => {
    try {
      setIsUserLoading(true);
      const user = await broadcastService.sp.web.currentUser();
      setCurrentUserId(user.Id);
    } catch (error) {
ErrorService.logError(
  "BroadcastTable.loadCurrentUser",
  error
);  
    } finally {
      setIsUserLoading(false);
    }
  };

  loadCurrentUser();
}, [broadcastService]);

  const [counts, setCounts] = useState({
    news: 0,
    events: 0,
    broadcasts: 0,
    brands: 0,
  });



  useEffect(() => {
    loadDashboardCounts();
  }, []);

  const loadDashboardCounts = async () => {
    const result = await broadcastService.getAllCounts();
    setCounts(result);
  };

  useEffect(() => {
    const url = window.location.pathname.toLowerCase();

    if (url.includes("cph_dashboard.aspx")) {
      setActiveMainTab("news");
    } else if (url.includes("cph_events.aspx")) {
      setActiveMainTab("events");
    } else if (url.includes("cph_broadcasts.aspx")) {
      setActiveMainTab("broadcasts");
    } else if (url.includes("cph_brands.aspx")) {
      setActiveMainTab("brands");
    }
  }, []);

  /* ---------------- FILTER DATA ---------------- */
  const formatRequestNumber = (item: any) => {
    const id = item?.Id || item?.ID;
    if (!id) return "-";

    const padded = ("00000000" + id.toString()).slice(-8);
    return `RQ${padded}`;
  };

  const matchesStatusTab = (item: any, tab: string, userId: number) => {
    const s = (item.Status || "").trim().toLowerCase();

    /* ===================== DRAFT ===================== */
    // All items created by logged in user which are saved but not submitted
    if (tab === "Draft") {
      return s === "draft" && item.AuthorId === userId;
    }

    /* ===================== PENDING ===================== */
    // All items pending WITH logged in user for approval
    if (tab === "Pending") {
      return (
        item.PendingWithId === userId &&
        (s === "pendingwithlm" ||
          s === "pendingwithcm" ||
          s === "pendingwithcreator")
      );
    }

    /* ===================== APPROVAL QUEUE ===================== */
    // All items created by logged in user and still pending (with LM or CM)
    if (tab === "Approval Queue") {
      return (
        item.AuthorId === userId &&
        (s === "pendingwithlm" || s === "pendingwithcm")
      );
    }

    /* ===================== PUBLISHED ===================== */
    // Published items where user is creator OR part of workflow
    if (tab === "Published") {
      return (
        s === "published" &&
        (item.AuthorId === userId ||
          item.AssignToId === userId ||
          item.FollowOnActionById === userId ||
          item.PendingWithId === userId)
      );
    }

    /* ===================== VIEW REQUESTS ===================== */
    // All items where user is involved — independent of status
    if (tab === "View Requests") {
      return (
        (item.AuthorId === userId &&
          (s === "draft" ||
            s === "pendingwithlm" ||
            s === "pendingwithcm" ||
            s === "pendingwithcreator" ||
            s === "published" ||
            s === "rejected")) ||
        (item.AssignToId === userId &&
          (s === "pendingwithlm" ||
            s === "pendingwithcm" ||
            s === "published" ||
            s === "rejected" ||
            s === "pendingwithcreator")) ||
        (item.FollowOnActionById === userId &&
          (s === "pendingwithcm" ||
            s === "published" ||
            s === "rejected" ||
            s === "pendingwithlm")) ||
        (item.PendingWithId === userId &&
          (s === "pendingwithlm" ||
            s === "pendingwithcm" ||
            s === "published" ||
            s === "rejected" ||
            s === "pendingwithcreator"))
      );
    }

    return false;
  };

    const isTableLoading = isLoading || isUserLoading;

  const filteredData = useMemo(() => {
            if (isTableLoading) return [];

    if (!items || !items.length) return [];

    return items.filter((item) => {
      /* ---------- STATUS FILTER ---------- */
      const matchesStatus = matchesStatusTab(
        item,
        selectedStatus,
        currentUserId,
      );

      /* ---------- SEARCH FILTER ---------- */
      const searchText = (search || "").trim().toLowerCase();

      const matchesSearch =
        !searchText ||
        (item.Title || "").toLowerCase().includes(searchText) ||
        formatRequestNumber(item).toLowerCase().includes(searchText);

      return matchesStatus && matchesSearch;
    });
  }, [items, selectedStatus, search, currentUserId, isTableLoading]);

  /* ================= IMAGE HELPER (ONE IMAGE ONLY) ================= */
  const getSingleImageUrl = (attachments: any[]) => {
    if (!attachments || attachments.length === 0) return null;

    const imageExtensions = [".png", ".jpg", ".jpeg", ".gif"];

    const image = attachments.find((att) =>
      imageExtensions.some((ext) => att.FileName?.toLowerCase().endsWith(ext)),
    );

    return image ? image.ServerRelativeUrl : null;
  };

  const getTabStatusName = (itemStatus: string) => {
    const s = (itemStatus || "").trim();

    switch (s.toLowerCase()) {
      case "draft":
        return "Draft";

      case "pendingwithcreator":
        return "PendingWithCreator";

      case "pendingwithlm":
        return "PendingWithLM";

      case "pendingwithcm":
        return "PendingWithCM";

      case "published":
        return "Published";

      case "rejected":
        return "Rejected";

      case "deleted":
        return "Deleted";

      case "closed":
        return "Closed";

      case "canceled":
        return "Canceled";

      default:
        return s;
    }
  };

  /* ---------------- PAGINATION ---------------- */
  const totalItems = filteredData.length;
  const totalPages = Math.ceil(totalItems / pageSize);
  const startIndex = (currentPage - 1) * pageSize;

  const pageItems = filteredData.slice(startIndex, startIndex + pageSize);

  React.useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages === 0 ? 1 : totalPages);
    }
  }, [pageSize, totalPages]);

  const formatISTDate = (dateValue: string) => {
  if (!dateValue) return "-";

  return new Date(dateValue).toLocaleDateString("en-GB", {
    timeZone: "Asia/Kolkata",
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}; 

  return (
    <div>
      <Container fluid className={styles.dashboardWrapper}>
        {/* Top Banner */}
        <div className={styles.topImageSection}>
          <div className={styles.topImageContent}>
            <Row className="my-2">
              <div className="d-flex justify-content-between">
                <h5 className={styles.pageHeadingWhite}>CPH Dashboard</h5>
                {isSiteAdmin && (
                  <Button
                    variant="outline-primary"
                    className={styles.navigateBtn}
                    onClick={() =>
                      window.open(
                        `${context.pageContext.web.absoluteUrl}/SitePages/CPHAdminDashboard.aspx`,
                        "_blank",
                      )
                    }
                  >
                    Admin <i className="bi bi-arrow-right" />
                  </Button>
                )}
              </div>
            </Row>

            {/* Dashboard Cards */}
            <Row className="g-3 justify-content-between">
              <Col xs={12} sm={6} md={4} lg={4} xl={3} xxl={3}>
                <Card className={styles.cardCommon}>
                  <Card.Body className="d-flex justify-content-between">
                    <div>
                      <Card.Subtitle className={styles.dashboardCardSubTitle}>
                        Total News
                      </Card.Subtitle>
                      <Card.Title className={styles.dashboardCardCount}>
                        {counts.news}
                      </Card.Title>
                    </div>
                    <img
                      alt="News"
                      src={newsIcon}
                      className={styles.dashboardCardIcon}
                    />
                  </Card.Body>
                </Card>
              </Col>

              <Col xs={12} sm={6} md={4} lg={4} xl={3} xxl={3}>
                <Card className={styles.cardCommon}>
                  <Card.Body className="d-flex justify-content-between">
                    <div>
                      <Card.Subtitle className={styles.dashboardCardSubTitle}>
                        Total Events
                      </Card.Subtitle>
                      <Card.Title className={styles.dashboardCardCount}>
                        {counts.events}
                      </Card.Title>
                    </div>
                    <img
                      alt="Events"
                      src={eventsIcon}
                      className={styles.dashboardCardIcon}
                    />
                  </Card.Body>
                </Card>
              </Col>

              <Col xs={12} sm={6} md={4} lg={4} xl={3} xxl={3}>
                <Card className={styles.cardCommon}>
                  <Card.Body className="d-flex justify-content-between">
                    <div>
                      <Card.Subtitle className={styles.dashboardCardSubTitle}>
                        Total Broadcasts
                      </Card.Subtitle>
                      <Card.Title className={styles.dashboardCardCount}>
                        {counts.broadcasts}
                      </Card.Title>
                    </div>
                    <img
                      alt="Broadcasts"
                      src={broadcastIcon}
                      className={styles.dashboardCardIcon}
                    />
                  </Card.Body>
                </Card>
              </Col>

              <Col xs={12} sm={6} md={4} lg={4} xl={3} xxl={3}>
                <Card className={styles.cardCommon}>
                  <Card.Body className="d-flex justify-content-between">
                    <div>
                      <Card.Subtitle className={styles.dashboardCardSubTitle}>
                        Total Brands
                      </Card.Subtitle>
                      <Card.Title className={styles.dashboardCardCount}>
                        {counts.brands}
                      </Card.Title>
                    </div>
                    <img
                      alt="Brands"
                      src={brandIcon}
                      className={styles.dashboardCardIcon}
                    />
                  </Card.Body>
                </Card>
              </Col>
            </Row>
          </div>
        </div>

        {/* Main Card */}
        <Row className="mt-4 mx-0 mx-md-1">
          <Col className="px-0 px-md-1">
            <Card className={`${styles.cardCommon}  p-0 p-md-2`}>
              <div className="card-body">
                {/* MAIN TABS */}
                <Tabs
                  activeKey={activeMainTab}
                  onSelect={(k) => {
                    if (!k) return;

                    if (k === "news") {
                      window.location.href = `${context.pageContext.web.absoluteUrl}/SitePages/CPH_Dashboard.aspx`;
                    }

                    if (k === "events") {
                      window.location.href = `${context.pageContext.web.absoluteUrl}/SitePages/Events.aspx`;
                    }

                    if (k === "broadcasts") {
                      window.location.href = `${context.pageContext.web.absoluteUrl}/SitePages/Broadcast.aspx`;
                    }

                    if (k === "brands") {
                      window.location.href = `${context.pageContext.web.absoluteUrl}/SitePages/Brand.aspx`;
                    }
                  }}
                  justify
                  className={styles.customTabs}
                >
                  {/* NEWS TAB */}
                  <Tab eventKey="news" title="News" />

                  {/* EVENTS TAB — MAIN CONTENT */}
                  <Tab eventKey="events" title="Events"></Tab>
                  <Tab eventKey="broadcasts" title="Broadcasts">
                    {/* STATUS + SEARCH + CREATE — SAME AS NEWS */}
                    <div className="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-4 mt-4 w-100">
                      {/* LEFT — STATUS */}
                      <ButtonGroup
                        className={`${styles.statusGroup} flex-wrap`}
                      >
                        {STATUSES.map((status) => (
                          <ToggleButton
                            key={status}
                            id={`status-${status}`}
                            type="radio"
                            variant="outline-primary"
                            name="statusFilter"
                            value={status}
                            checked={selectedStatus === status}
                            onChange={(e) => {
                              onStatusChange(e.currentTarget.value);
                              setCurrentPage(1);
                            }}
                            className={styles.statusBtn}
                          >
                            {status}
                          </ToggleButton>
                        ))}
                      </ButtonGroup>

                      {/* RIGHT — SEARCH + CREATE */}
                      <div
                        className="d-flex align-items-center gap-2 ms-auto"
                        style={{ height: 44 }}
                      >
                        {/* SEARCH */}
                        <div className={`input-group ${styles.searchBox}`}>
                          <span className="input-group-text bg-white border-end-0">
                            <i className="bi bi-search" />
                          </span>

                          <Form.Control
                            type="text"
                            placeholder="Search..."
                            value={search}
                            onChange={(e) => {
                              onSearchChange(e.target.value);
                              setCurrentPage(1);
                            }}
                            className={`border-start-0 ${styles.noFocus}`}
                          />
                        </div>

                        {/* CREATE EVENT */}
                        {canCreateBroadcast && (
                          <Button
                            className={`mt-3 ${styles.createNewsBtn}`}
                            onClick={onCreate}
                          >
                            <i className="bi bi-plus-circle me-1" />
                            Create Broadcast
                          </Button>
                        )}
                      </div>
                    </div>

                    {/* TABLE */}
                    <Table responsive className={styles.table}>
                      <thead>
                        <tr>
                          <th>Request Number</th>
                          <th>Image</th>
                          <th>Title</th>
                          <th>Broadcast Date</th>
                          <th>Created On</th>
                          <th>Created By</th>
                          <th>Pending With</th>
                          <th>Status</th>
                          <th>Action</th>
                        </tr>
                      </thead>

        <tbody>
  {isTableLoading ? (
    <tr>
      <td colSpan={9} className="text-center py-4">
        <div className="d-flex justify-content-center align-items-center gap-2">
          <div
            className="spinner-border spinner-border-sm text-primary"
            role="status"
          />
          <span>Loading data...</span>
        </div>
      </td>
    </tr>
  ) : pageItems.length === 0 ? (
    <tr>
      <td colSpan={9} className="text-center py-4">
        No records found
      </td>
    </tr>
  ) : (
    pageItems.map((item, index) => (
      <tr key={item.Id || index}>
        <td>
          <a
            href="#"
            className={styles.documentLink}
            onClick={(e) => {
              e.preventDefault();
              onPreviewPopup(item);
            }}
          >
            {formatRequestNumber(item)}
          </a>
        </td>

        <td>
          {getSingleImageUrl(item.AttachmentsData) ? (
            <img
              src={getSingleImageUrl(item.AttachmentsData)!}
              className={styles.broadcastImage}
              alt="Broadcast"
            />
          ) : (
            <span className={styles.noImage}>-</span>
          )}
        </td>
 
        <td className={styles.titleCol}>{item.Title}</td>

        <td>{formatISTDate(item.EventDate)}</td>
        <td>{formatISTDate(item.Created)}</td>

        <td>{item.Author?.Title}</td>
        <td>{item.PendingWith?.Title}</td>
        <td>{getTabStatusName(item.Status)}</td>

        <td className={styles.tableIcon}>
          <i
            className="bi bi-pencil-square"
            onClick={() => onEdit(item)}
            style={{ cursor: "pointer" }}
          />
        </td>
      </tr>
    ))
  )}
</tbody> 
                    </Table>

                    {/* FOOTER */}
                    <div className={styles.tableFooter}>
                      <div className={styles.entriesInfo}>
                        Showing {totalItems === 0 ? 0 : startIndex + 1}–
                        {Math.min(startIndex + pageSize, totalItems)} of{" "}
                        {totalItems} entries
                      </div>

                      {/*  RIGHT SIDE GROUP */}
                      <div className="d-flex align-items-center gap-3">
                        {/* Entries per page */}
                        <div className="d-flex align-items-center gap-2">
                          <Form.Select
                            style={{ width: "90px" }}
                            value={pageSize}
                            onChange={(e) => {
                              setPageSize(Number(e.target.value));
                              setCurrentPage(1);
                            }}
                          >
                            {PAGE_SIZES.map((n) => (
                              <option key={n} value={n}>
                                {n}
                              </option>
                            ))}
                          </Form.Select>
                          <span className="mb-0">entries per page</span>
                        </div>

                        {/* Pagination */}
                        <Pagination
                          size="sm"
                          className={`${styles.customPagination} d-flex align-items-center mb-0`}
                        >
                          <Pagination.Prev
                            disabled={currentPage === 1}
                            onClick={() => setCurrentPage((p) => p - 1)}
                          />

                          {[...Array(totalPages)].map((_, i) => (
                            <Pagination.Item
                              key={i}
                              active={currentPage === i + 1}
                              onClick={() => setCurrentPage(i + 1)}
                            >
                              {i + 1}
                            </Pagination.Item>
                          ))}

                          <Pagination.Next
                            disabled={
                              currentPage === totalPages || totalPages === 0
                            }
                            onClick={() => setCurrentPage((p) => p + 1)}
                          />
                        </Pagination>
                      </div>
                    </div>
                  </Tab>
                  <Tab eventKey="brands" title="Brands"></Tab>
                </Tabs>
              </div>
            </Card>
          </Col>
        </Row>
      </Container>
    </div>
  );
};
 
export default BroadcastTable;
