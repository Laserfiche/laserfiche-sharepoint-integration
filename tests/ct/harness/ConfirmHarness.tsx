import * as React from 'react';
import { useConfirm } from '../../../src/extensions/savetoLaserfiche/CommonDialogs';

declare global {
  interface Window {
    __confirmResults: Array<boolean | 'pending'>;
  }
}

export default function ConfirmHarness(props: {
  label: string;
  cancelButtonText: string;
  headerText: string;
}): JSX.Element {
  const [getConfirmation, Confirmation] = useConfirm();

  const ask = async (): Promise<void> => {
    window.__confirmResults = window.__confirmResults ?? [];
    window.__confirmResults.push('pending');
    const result = await getConfirmation(props.label);
    window.__confirmResults[window.__confirmResults.length - 1] = result as boolean;
  };

  return (
    <>
      <button onClick={ask}>Ask</button>
      <Confirmation cancelButtonText={props.cancelButtonText} headerText={props.headerText} />
    </>
  );
}
