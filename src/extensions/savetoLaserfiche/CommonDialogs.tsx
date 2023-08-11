import * as React from 'react';
import styles from './SendToLaserFiche.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SavedToLaserficheDocumentData } from './SaveDocumentToLaserfiche';

const SAVING_DOCUMENT_TO_LASERFICHE = 'Saving document to Laserfiche...';

export default function LoadingDialog(): JSX.Element {
  return (
    <>
      <img src='/_layouts/15/images/progress.gif' />
      <br />
      <div>{SAVING_DOCUMENT_TO_LASERFICHE}</div>
    </>
  );
}

const DOCUMENT_SUCCESSFULLY_UPLOADED_TO_LASERFICHE_WITH_NAME =
  'Document successfully uploaded to Laserfiche with name:';
const METADATA_FAILED_TO_SAVE_INVALID_FIELD =
  'All metadata failed to save due to at least one invalid field';
const TEMPLATE_FIELDS_NOT_APPLIED =
  'The Laserfiche template and fields were not applied to this document.';
const CLOSE = 'Close';
const VIEW_FILE_IN_LASERFICHE = 'View file in Laserfiche';

export function SavedToLaserficheSuccessDialogText(props: {
  successfulSave: SavedToLaserficheDocumentData;
}): JSX.Element {
  React.useEffect(() => {
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
    );
  }, []);

  const metadataFailedNotice: JSX.Element = (
    <span>
      {METADATA_FAILED_TO_SAVE_INVALID_FIELD}
      <br /> <p style={{ color: 'red' }}>{TEMPLATE_FIELDS_NOT_APPLIED}</p>
    </span>
  );

  return (
    <>
      <div>
        {`${DOCUMENT_SUCCESSFULLY_UPLOADED_TO_LASERFICHE_WITH_NAME}  ${props.successfulSave.fileName}.`}
        {!props.successfulSave.metadataSaved && metadataFailedNotice}
      </div>
    </>
  );
}

export function SavedToLaserficheSuccessDialogButtons(props: {
  closeClick: () => Promise<void>;
  successfulSave: SavedToLaserficheDocumentData;
}): JSX.Element {
  React.useEffect(() => {
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
    );
  }, []);

  function viewFile(): void {
    window.open(props.successfulSave.fileLink);
  }

  return (
    <>
      {props.successfulSave?.fileLink && (
        <button
          className={`lf-button primary-button ${styles.actionButton}`}
          title={VIEW_FILE_IN_LASERFICHE}
          onClick={viewFile}
        >
          {VIEW_FILE_IN_LASERFICHE}
        </button>
      )}
      <button className='lf-button sec-button' onClick={props.closeClick}>
        {CLOSE}
      </button>
    </>
  );
}
