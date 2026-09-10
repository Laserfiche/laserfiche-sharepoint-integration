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
const CLOSE = 'Close';
const VIEW_FILE_IN_LASERFICHE = 'View file in Laserfiche';

export function SavedToLaserficheSuccessDialogText(props: {
  successfulSave: SavedToLaserficheDocumentData;
}): JSX.Element {
  React.useEffect(() => {
    SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
    SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
  }, []);

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
                  src='data:image/png;base64,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'
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
