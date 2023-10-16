import * as React from 'react';
import styles from './SendToLaserFiche.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SavedToLaserficheDocumentData } from './SaveDocumentToLaserfiche';
import {
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
} from '../../webparts/constants';
import { ActionTypes } from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';

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
const EXISTING_SP_DOCUMENT_DELETED =
  'The existing SharePoint document was deleted.';
const EXISTING_SP_DOCUMENT_REPLACED =
  'The existing SharePoint document was replaced with a link to the document in Laserfiche.';
const METADATA_FAILED_TO_SAVE_INVALID_FIELD =
  'Unable to save metadata due to at least one invalid field value.';
const CLOSE = 'Close';
const VIEW_FILE_IN_LASERFICHE = 'View file in Laserfiche';

const WARNING = 'Warning: ';
const ERROR_DETAILS = 'Error details';
export function SavedToLaserficheSuccessDialogText(props: {
  successfulSave: SavedToLaserficheDocumentData;
}): JSX.Element {
  React.useEffect(() => {
    SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
    SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
  }, []);

  const metadataFailedNotice: JSX.Element = (
    <>
      <div className={styles.paddingUnder}>
        <b>{WARNING}</b>
        {METADATA_FAILED_TO_SAVE_INVALID_FIELD}
      </div>
      <Collapsible title={ERROR_DETAILS}>
        {props.successfulSave.failedMetadata}
      </Collapsible>
    </>
  );

  return (
    <>
      <div className={styles.successSaveToLaserfiche}>
        <div className={styles.paddingUnder}>
          {`${DOCUMENT_SUCCESSFULLY_UPLOADED_TO_LASERFICHE_WITH_NAME}  ${props.successfulSave.fileName}.`}
        </div>
        <div>
          {props.successfulSave.action === ActionTypes.MOVE_AND_DELETE &&
            EXISTING_SP_DOCUMENT_DELETED}
          {props.successfulSave.action === ActionTypes.REPLACE &&
            EXISTING_SP_DOCUMENT_REPLACED}
        </div>
        {!props.successfulSave.metadataSaved && metadataFailedNotice}
      </div>
    </>
  );
}

export function Collapsible(props: {
  open?: boolean;
  children: JSX.Element;
  title: string;
}): JSX.Element {
  const [isOpen, setIsOpen] = React.useState<boolean>(props.open ?? false);

  const handleFilterOpening: () => void = () => {
    setIsOpen((prev) => !prev);
  };

  return (
    <>
      <div className={styles.collapseBox}>
        <button
          className={styles.lfMaterialIconButton}
          onClick={handleFilterOpening}
        >
          {!isOpen ? (
            <span className='material-icons-outlined'> chevron_right </span>
          ) : (
            <span className='material-icons-outlined'> expand_less </span>
          )}
        </button>
        <span>{props.title}</span>
      </div>

      {isOpen && props.children}
    </>
  );
}

export function SavedToLaserficheSuccessDialogButtons(props: {
  closeClick: () => Promise<void>;
  successfulSave: SavedToLaserficheDocumentData;
}): JSX.Element {
  React.useEffect(() => {
    SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
    SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
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

const createPromise: () => Promise<boolean>[] = () => {
  let resolver;
  return [
    new Promise<boolean>((resolve, reject) => {
      resolver = resolve;
    }),
    resolver,
  ];
};

export const useConfirm: () => [
  (text: string) => Promise<unknown>,
  (props: { cancelButtonText: string }) => JSX.Element
] = () => {
  const [open, setOpen] = React.useState(false);
  const [resolver, setResolver] = React.useState({ resolve: null });
  const [label, setLabel] = React.useState('');

  const getConfirmation: (text: string) => Promise<boolean> = async (text: string) => {
    setLabel(text);
    setOpen(true);
    const [promise, resolve] = await createPromise();
    setResolver({ resolve: resolve });
    return promise;
  };

  const onClick: (status: boolean) => Promise<void> = async (status: boolean) => {
    setOpen(false);
    resolver.resolve(status);
  };

  const Confirmation: (props: {
    cancelButtonText: string;
  }) => JSX.Element = (props: { cancelButtonText: string }) => (
    <>
      {open && (
        <>
          <div className={`modal-header ${styles.header}`}>
            <div className='modal-title' id='ModalLabel'>
              <div className={styles.logoHeader}>
                <img
                  src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII='
                  width='30'
                  height='30'
                />
                <span className={styles.paddingLeft}>
                  Document already exists
                </span>
              </div>
            </div>
          </div>
          <div className={`modal-body ${styles.contentBox}`}>
            <span>{label}</span>
          </div>
          <div className={`modal-footer ${styles.footer}`}>
            <button
              className={`lf-button primary-button ${styles.actionButton}`}
              onClick={() => onClick(true)}
            >
              Continue
            </button>
            <button
              className='lf-button sec-button'
              onClick={() => onClick(false)}
            >
              {props.cancelButtonText}
            </button>
          </div>
        </>
      )}
    </>
  );

  return [getConfirmation, Confirmation];
};
