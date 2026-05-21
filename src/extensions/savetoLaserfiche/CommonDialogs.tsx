// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import styles from './SendToLaserFiche.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SavedToLaserficheDocumentData } from './SaveDocumentToLaserfiche';
import {
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
} from '../../webparts/constants';
import { ActionTypes } from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import { CONTINUE } from '../../webparts/strings';

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

export function MessageDialog(props: {
  title: string;
  message: string;
  clickOkay: () => void;
}): JSX.Element {
  return (
    <div className='modal-dialog modal-dialog-centered'>
      <div className={`modal-content ${styles.wrapper}`}>
        <div className={styles.header}>
          <h5 className='modal-title' id='ModalLabel'>
            {props.title}
          </h5>
        </div>
        <div className={styles.contentBox}>{props.message}</div>
        <div className={styles.footer}>
          <button
            type='button'
            className='lf-button primary-button'
            data-dismiss='modal'
            onClick={props.clickOkay}
          >
            Okay
          </button>
        </div>
      </div>
    </div>
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
  (props: { cancelButtonText: string; headerText: string }) => JSX.Element
] = () => {
  const [open, setOpen] = React.useState(false);
  const [resolver, setResolver] = React.useState({ resolve: null });
  const [label, setLabel] = React.useState('');

  const getConfirmation: (text: string) => Promise<boolean> = async (
    text: string
  ) => {
    setLabel(text);
    setOpen(true);
    const [promise, resolve] = await createPromise();
    setResolver({ resolve: resolve });
    return promise;
  };

  const onClick: (status: boolean) => Promise<void> = async (
    status: boolean
  ) => {
    setOpen(false);
    resolver.resolve(status);
  };

  const Confirmation: (props: {
    cancelButtonText: string;
    headerText: string;
  }) => JSX.Element = (props: {
    cancelButtonText: string;
    headerText: string;
  }) => (
    <>
      {open && (
        <>
          <div className={`modal-header ${styles.header}`}>
            <div className='modal-title' id='ModalLabel'>
              <div className={styles.logoHeader}>
                <img
                  src={require('../../../sharepoint/assets/laserfiche-logo.png')}
                  width='30'
                  height='30'
                />
                <span className={styles.paddingLeft}>{props.headerText}</span>
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
              {CONTINUE}
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
