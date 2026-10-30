import * as React from "react";
import {Modal, Carousel } from "react-bootstrap";
import styles from "./MyViewBroadcast.module.scss";

    
interface Props {
  show: boolean;
  onClose: () => void;
  item: any;
  attachments?: { FileName: string; ServerRelativeUrl: string }[];
}  
 
const PreviewBroadcastModal: React.FC<Props> = ({
  show,
  onClose,
  item,
  attachments = [],
}) => {
  if (!item) return null;
  

  // Helpers
const isImage = (fileName: string) =>
  /\.(png|jpe?g|gif|bmp)$/i.test(fileName);

const isPdf = (fileName: string) =>
  /\.pdf$/i.test(fileName);

const isWord = (fileName: string) =>
  /\.(doc|docx)$/i.test(fileName);
 
  // Separate attachments
  const imageAttachments = attachments.filter(a =>
    isImage(a.FileName)
  );

  const pdfAttachments = attachments.filter(a =>
    isPdf(a.FileName)
  ); 


const wordAttachments = attachments.filter(a =>
  isWord(a.FileName)
);

const openRelatedSite = () => {
  const url =
    typeof item?.RelatedSiteURL === "string"
      ? item.RelatedSiteURL
      : item?.RelatedSiteURL?.Url;

  if (url) {
    window.open(url, "_blank");
  }
};  

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
    <Modal show={show} onHide={onClose} size="lg" centered> 
      <Modal.Header closeButton>
        <Modal.Title>{item.Title}</Modal.Title> 
      </Modal.Header>
 
      <Modal.Body className={`p-3 ${styles.modalBody}`}>
        {/* ================= IMAGES ================= */}
 

        {imageAttachments.length === 1 && (
  <img
    src={imageAttachments[0].ServerRelativeUrl}
    alt="Broadcast"
    onClick={openRelatedSite} 
    style={{
      width: "100%",
                        height: "auto",
      objectFit: "cover",
      marginBottom: "16px",
      cursor: "pointer"
    }}
  />
)} 
  

        {/* Multiple Images → Carousel */}
{imageAttachments.length > 1 && ( 
 <Carousel
    className={`${styles.modalCarousel} mb-3`}
    prevIcon={
      <span
        aria-hidden="true"
        className="carousel-control-prev-icon"
        style={{ filter: "invert(1)" }}
      />
    }
    nextIcon={
      <span
        aria-hidden="true"
        className="carousel-control-next-icon"
        style={{ filter: "invert(1)" }}
      />
    }
  > 
    {imageAttachments.map((img, index) => (
      <Carousel.Item key={index}>
        <img
          src={img.ServerRelativeUrl}
          alt={img.FileName}
          onClick={index === 0 ? openRelatedSite : undefined}
          style={{
            width: "100%",
            height: "auto",
            objectFit: "cover",
            cursor: index === 0 ? "pointer" : "default"
          }}
        />
      </Carousel.Item>
    ))}
  </Carousel>
)}  
  
        {/* ================= EVENT DATE ================= */}
        <p className="mb-3 text-muted">
           {formatISTDate(item.EventDate)}

        </p>
 
        
{(pdfAttachments.length > 0 || wordAttachments.length > 0) && (
  <div className="mt-3">
    <h6>Attachments</h6>
    <ul className="list-unstyled">

      {/* PDF Files */}
      {pdfAttachments.map((pdf, index) => (
                <li key={index}>
                  <a
                    href={pdf.ServerRelativeUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    📄 {pdf.FileName}
                  </a>
                </li>
              ))}

      {/* Word Files */}
      {wordAttachments.map((file, index) => (
        <li key={`word-${index}`}>
          <a
            href={file.ServerRelativeUrl}
            target="_blank"
            rel="noopener noreferrer"
          >
            📝 {file.FileName}
          </a> 
        </li>
      ))}

    </ul>
  </div>
)}     
      </Modal.Body>
 
      {/* <Modal.Footer>
        <Button variant="outline-primary" onClick={onClose}>
          Close
        </Button>
      </Modal.Footer> */}
    </Modal>
  );
};
 
export default PreviewBroadcastModal;  