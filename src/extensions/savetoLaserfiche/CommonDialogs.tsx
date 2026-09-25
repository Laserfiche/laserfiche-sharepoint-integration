// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import styles from './SendToLaserFiche.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SavedLaserficheDocument } from '../../Utils/Types';
import {
  LASERFICHE_ICON_URL,
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
  LOADING_SPINNER_DELAY_MS,
} from '../../webparts/constants';
import { ActionTypes } from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import {
  CLOSE,
  CONTINUE,
  LASERFICHE,
  SAVED_A_COPY_TO_LASERFICHE,
  SHOW_IN_FOLDER,
} from '../../webparts/strings';

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

export function DelayedSpinner(props: {
  loading: boolean;
  label: string;
  delayMs?: number;
  className?: string;
}): JSX.Element {
  const [visible, setVisible] = React.useState<boolean>(false);
  const delayMs = props.delayMs ?? LOADING_SPINNER_DELAY_MS;

  React.useEffect(() => {
    if (!props.loading) {
      setVisible(false);
      return;
    }
    const timeout = setTimeout(() => setVisible(true), delayMs);
    return () => clearTimeout(timeout);
  }, [props.loading, delayMs]);

  if (!visible) {
    return null;
  }
  return (
    <div
      role='status'
      className={`d-flex align-items-center gap-2${
        props.className ? ` ${props.className}` : ''
      }`}
    >
      <span className='spinner-border spinner-border-sm' aria-hidden='true' />
      <span>{props.label}</span>
    </div>
  );
}

const EXISTING_SP_DOCUMENT_DELETED =
  'The existing SharePoint document was deleted.';
const EXISTING_SP_DOCUMENT_REPLACED =
  'The existing SharePoint document was replaced with a link to the document in Laserfiche.';

export function SavedToLaserficheSuccessDialogText(props: {
  successfulSave: SavedLaserficheDocument;
  action?: ActionTypes;
}): JSX.Element {
  React.useEffect(() => {
    SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
    SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
  }, []);

  const { fileName, fileLink, folderLink } = props.successfulSave;
  return (
    <div className={styles.successSaveToLaserfiche}>
      <div className={styles.paddingUnder}>{SAVED_A_COPY_TO_LASERFICHE}</div>
      <ul className={`${styles.noMargin} ${styles.paddingUnder}`}>
        <li>
          {fileLink ? (
            <a href={fileLink} target='_blank' rel='noreferrer'>
              {fileName}
            </a>
          ) : (
            fileName
          )}
        </li>
      </ul>
      {folderLink && (
        <div className={styles.paddingUnder}>
          <a href={folderLink} target='_blank' rel='noreferrer'>
            {SHOW_IN_FOLDER}
          </a>
        </div>
      )}
      <div>
        {props.action === ActionTypes.MOVE_AND_DELETE &&
          EXISTING_SP_DOCUMENT_DELETED}
        {props.action === ActionTypes.REPLACE && EXISTING_SP_DOCUMENT_REPLACED}
      </div>
    </div>
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
}): JSX.Element {
  React.useEffect(() => {
    SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
    SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
  }, []);

  return (
    <button className='lf-button sec-button' onClick={props.closeClick}>
      {CLOSE}
    </button>
  );
}

export function SavedToLaserficheSuccessDialog(props: {
  successfulSave: SavedLaserficheDocument;
  closeClick: () => Promise<void>;
}): JSX.Element {
  return (
    <div className='modal-dialog modal-dialog-centered'>
      <div className={`modal-content ${styles.wrapper}`}>
        <div className={styles.header}>
          <LaserficheDialogTitle title={LASERFICHE} />
        </div>
        <div className={styles.contentBox}>
          <SavedToLaserficheSuccessDialogText
            successfulSave={props.successfulSave}
          />
        </div>
        <div className={styles.footer}>
          <SavedToLaserficheSuccessDialogButtons
            closeClick={props.closeClick}
          />
        </div>
      </div>
    </div>
  );
}

export function LaserficheDialogTitle(props: { title: string }): JSX.Element {
  return (
    <div className={styles.logoHeader}>
      <img src={LASERFICHE_ICON_URL} alt='' width='30' height='30' />
      <span className={styles.paddingLeft}>{props.title}</span>
    </div>
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
              <LaserficheDialogTitle title={props.headerText} />
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
