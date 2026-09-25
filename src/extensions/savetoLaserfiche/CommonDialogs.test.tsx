// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import { act, fireEvent, render, screen, waitFor } from '@testing-library/react';
import LoadingDialog, {
  DelayedSpinner,
  SavedToLaserficheSuccessDialog,
  SavedToLaserficheSuccessDialogText,
  Collapsible,
  SavedToLaserficheSuccessDialogButtons,
  MessageDialog,
  LaserficheDialogTitle,
  useConfirm,
} from './CommonDialogs';
import { ActionTypes } from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import { LASERFICHE_ICON_URL } from '../../webparts/constants';
import {
  CLOSE,
  CONTINUE,
  LASERFICHE,
  SAVED_A_COPY_TO_LASERFICHE,
  SHOW_IN_FOLDER,
} from '../../webparts/strings';
import type { SavedLaserficheDocument } from '../../Utils/Types';

describe('LoadingDialog', () => {
  test('renders loading text and progress image', () => {
    render(<LoadingDialog />);

    expect(
      screen.getByText('Saving document to Laserfiche...')
    ).toBeInTheDocument();

    const img = screen.getByRole('img') as HTMLImageElement;
    expect(img).toHaveAttribute('src', '/_layouts/15/images/progress.gif');
  });
});

describe('DelayedSpinner', () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  function advance(ms: number): void {
    act(() => {
      vi.advanceTimersByTime(ms);
    });
  }

  test('appears only once loading has lasted 500ms', () => {
    render(<DelayedSpinner loading={true} label='Loading things...' />);

    advance(499);
    expect(screen.queryByRole('status')).not.toBeInTheDocument();

    advance(1);
    expect(screen.getByRole('status')).toHaveTextContent('Loading things...');
  });

  test('disappears as soon as loading ends', () => {
    const { rerender } = render(
      <DelayedSpinner loading={true} label='Loading things...' />
    );
    advance(500);
    expect(screen.getByRole('status')).toBeInTheDocument();

    rerender(<DelayedSpinner loading={false} label='Loading things...' />);

    expect(screen.queryByRole('status')).not.toBeInTheDocument();
  });

  test('never appears when loading ends within 500ms', () => {
    const { rerender } = render(
      <DelayedSpinner loading={true} label='Loading things...' />
    );
    advance(300);

    rerender(<DelayedSpinner loading={false} label='Loading things...' />);
    advance(1000);

    expect(screen.queryByRole('status')).not.toBeInTheDocument();
  });

  test('restarts the delay when loading starts again', () => {
    const { rerender } = render(
      <DelayedSpinner loading={true} label='Loading things...' />
    );
    advance(500);
    rerender(<DelayedSpinner loading={false} label='Loading things...' />);

    rerender(<DelayedSpinner loading={true} label='Loading things...' />);
    advance(499);
    expect(screen.queryByRole('status')).not.toBeInTheDocument();

    advance(1);
    expect(screen.getByRole('status')).toBeInTheDocument();
  });

  test('applies the given className to the status row', () => {
    render(
      <DelayedSpinner
        loading={true}
        label='Loading things...'
        className='ms-2'
      />
    );

    advance(500);

    expect(screen.getByRole('status')).toHaveClass('ms-2');
  });

  test('unmounting cancels the pending delay', () => {
    const { unmount } = render(
      <DelayedSpinner loading={true} label='Loading things...' />
    );

    unmount();

    expect(vi.getTimerCount()).toBe(0);
  });
});

const FILE_LINK = 'https://webclient.example.com/DocView.aspx?id=42';
const FOLDER_LINK = 'https://webclient.example.com/Browse.aspx#?id=7';

function buildSavedDocument(
  overrides: Partial<SavedLaserficheDocument> = {}
): SavedLaserficheDocument {
  return {
    fileName: 'Contract',
    fileLink: FILE_LINK,
    folderLink: FOLDER_LINK,
    ...overrides,
  };
}

describe('SavedToLaserficheSuccessDialogText', () => {
  test('says a copy was saved and how to keep editing', () => {
    render(
      <SavedToLaserficheSuccessDialogText
        successfulSave={buildSavedDocument()}
      />
    );

    expect(screen.getByText(SAVED_A_COPY_TO_LASERFICHE)).toBeInTheDocument();
    expect(
      screen.queryByText(/successfully uploaded/i)
    ).not.toBeInTheDocument();
  });

  test('links the document name to the document in a new tab', () => {
    render(
      <SavedToLaserficheSuccessDialogText
        successfulSave={buildSavedDocument()}
      />
    );

    const documentLink = screen.getByRole('link', { name: 'Contract' });
    expect(documentLink).toHaveAttribute('href', FILE_LINK);
    expect(documentLink).toHaveAttribute('target', '_blank');
    expect(documentLink.closest('li')).toBeInTheDocument();
  });

  test('links Show in folder to the parent folder in a new tab', () => {
    render(
      <SavedToLaserficheSuccessDialogText
        successfulSave={buildSavedDocument()}
      />
    );

    const folderLink = screen.getByRole('link', { name: SHOW_IN_FOLDER });
    expect(folderLink).toHaveAttribute('href', FOLDER_LINK);
    expect(folderLink).toHaveAttribute('target', '_blank');
  });

  test('shows the document name as plain text when there is no document link', () => {
    render(
      <SavedToLaserficheSuccessDialogText
        successfulSave={buildSavedDocument({ fileLink: undefined })}
      />
    );

    expect(screen.getByText('Contract')).toBeInTheDocument();
    expect(
      screen.queryByRole('link', { name: 'Contract' })
    ).not.toBeInTheDocument();
  });

  test('leaves out Show in folder when there is no folder link', () => {
    render(
      <SavedToLaserficheSuccessDialogText
        successfulSave={buildSavedDocument({ folderLink: undefined })}
      />
    );

    expect(screen.queryByText(SHOW_IN_FOLDER)).not.toBeInTheDocument();
  });

  test('renders SharePoint document deleted message for MOVE_AND_DELETE action', () => {
    render(
      <SavedToLaserficheSuccessDialogText
        successfulSave={buildSavedDocument()}
        action={ActionTypes.MOVE_AND_DELETE}
      />
    );

    expect(
      screen.getByText('The existing SharePoint document was deleted.')
    ).toBeInTheDocument();
    expect(screen.queryByText(/replaced with a link/i)).not.toBeInTheDocument();
  });

  test('renders SharePoint document replaced message for REPLACE action', () => {
    render(
      <SavedToLaserficheSuccessDialogText
        successfulSave={buildSavedDocument()}
        action={ActionTypes.REPLACE}
      />
    );

    expect(screen.getByText(/replaced with a link/i)).toBeInTheDocument();
    expect(screen.queryByText(/was deleted\./i)).not.toBeInTheDocument();
  });

  test('renders neither extra message for COPY action', () => {
    render(
      <SavedToLaserficheSuccessDialogText
        successfulSave={buildSavedDocument()}
        action={ActionTypes.COPY}
      />
    );

    expect(screen.queryByText(/was deleted\./i)).not.toBeInTheDocument();
    expect(screen.queryByText(/replaced with a link/i)).not.toBeInTheDocument();
  });
});

describe('Collapsible', () => {
  test('renders collapsed by default, hiding children and showing chevron_right', () => {
    render(
      <Collapsible title='My Section'>
        <div>Child Content</div>
      </Collapsible>
    );

    expect(screen.queryByText('Child Content')).not.toBeInTheDocument();
    expect(screen.getByText('chevron_right')).toBeInTheDocument();
    expect(screen.queryByText('expand_less')).not.toBeInTheDocument();
  });

  test('renders expanded when open is true, showing children and expand_less', () => {
    render(
      <Collapsible title='My Section' open={true}>
        <div>Child Content</div>
      </Collapsible>
    );

    expect(screen.getByText('Child Content')).toBeInTheDocument();
    expect(screen.getByText('expand_less')).toBeInTheDocument();
    expect(screen.queryByText('chevron_right')).not.toBeInTheDocument();
  });

  test('clicking the toggle button flips child visibility and icon', () => {
    render(
      <Collapsible title='My Section'>
        <div>Child Content</div>
      </Collapsible>
    );

    const toggleButton = screen.getByRole('button');

    fireEvent.click(toggleButton);

    expect(screen.getByText('Child Content')).toBeInTheDocument();
    expect(screen.getByText('expand_less')).toBeInTheDocument();
    expect(screen.queryByText('chevron_right')).not.toBeInTheDocument();

    fireEvent.click(toggleButton);

    expect(screen.queryByText('Child Content')).not.toBeInTheDocument();
    expect(screen.getByText('chevron_right')).toBeInTheDocument();
    expect(screen.queryByText('expand_less')).not.toBeInTheDocument();
  });
});

describe('SavedToLaserficheSuccessDialogButtons', () => {
  test('renders only a Close button, which invokes closeClick', () => {
    const closeClick = vi.fn().mockResolvedValue(undefined);

    render(<SavedToLaserficheSuccessDialogButtons closeClick={closeClick} />);

    expect(screen.getAllByRole('button')).toHaveLength(1);
    expect(
      screen.queryByText('View file in Laserfiche')
    ).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole('button', { name: CLOSE }));
    expect(closeClick).toHaveBeenCalledTimes(1);
  });
});

describe('SavedToLaserficheSuccessDialog', () => {
  test('shows the Laserfiche title, the saved-copy message and both links', () => {
    render(
      <SavedToLaserficheSuccessDialog
        successfulSave={buildSavedDocument()}
        closeClick={vi.fn()}
      />
    );

    expect(screen.getByText(LASERFICHE)).toBeInTheDocument();
    expect(screen.getByRole('img')).toHaveAttribute('src', LASERFICHE_ICON_URL);
    expect(screen.getByText(SAVED_A_COPY_TO_LASERFICHE)).toBeInTheDocument();
    expect(screen.getByRole('link', { name: 'Contract' })).toHaveAttribute(
      'href',
      FILE_LINK
    );
    expect(screen.getByRole('link', { name: SHOW_IN_FOLDER })).toHaveAttribute(
      'href',
      FOLDER_LINK
    );
  });

  test('Close invokes closeClick', () => {
    const closeClick = vi.fn().mockResolvedValue(undefined);
    render(
      <SavedToLaserficheSuccessDialog
        successfulSave={buildSavedDocument()}
        closeClick={closeClick}
      />
    );

    fireEvent.click(screen.getByRole('button', { name: CLOSE }));

    expect(closeClick).toHaveBeenCalledTimes(1);
  });
});

describe('MessageDialog', () => {
  test('renders title and message, and Okay invokes clickOkay', () => {
    const clickOkay = vi.fn();

    render(
      <MessageDialog
        title='My Title'
        message='My Message'
        clickOkay={clickOkay}
      />
    );

    expect(screen.getByText('My Title')).toBeInTheDocument();
    expect(screen.getByText('My Message')).toBeInTheDocument();

    fireEvent.click(screen.getByText('Okay'));
    expect(clickOkay).toHaveBeenCalled();
  });
});

describe('LaserficheDialogTitle', () => {
  test('renders the title beside the decorative Laserfiche icon', () => {
    render(<LaserficheDialogTitle title='My Dialog' />);

    expect(screen.getByText('My Dialog')).toBeInTheDocument();
    expect(screen.getByRole('img')).toHaveAttribute(
      'src',
      LASERFICHE_ICON_URL
    );
  });
});

describe('useConfirm', () => {
  function ConfirmHost(props: { onResult: (r: boolean) => void }): JSX.Element {
    const [getConfirmation, Confirmation] = useConfirm();
    return (
      <>
        <button
          onClick={async () => props.onResult((await getConfirmation('Are you sure?')) as boolean)}
        >
          Ask
        </button>
        <Confirmation cancelButtonText='Go back' headerText='Please Confirm' />
      </>
    );
  }

  test('confirm UI is not shown until getConfirmation is invoked', () => {
    render(<ConfirmHost onResult={vi.fn()} />);

    expect(screen.queryByText('Please Confirm')).not.toBeInTheDocument();
    expect(screen.queryByText('Are you sure?')).not.toBeInTheDocument();
  });

  test('shows the header text beside the Laserfiche icon', async () => {
    render(<ConfirmHost onResult={vi.fn()} />);

    await act(async () => {
      fireEvent.click(screen.getByText('Ask'));
    });

    expect(screen.getByText('Please Confirm')).toBeInTheDocument();
    expect(screen.getByRole('img')).toHaveAttribute(
      'src',
      LASERFICHE_ICON_URL
    );
  });

  test('clicking Continue resolves true and hides the confirm UI', async () => {
    const onResult = vi.fn();
    render(<ConfirmHost onResult={onResult} />);

    await act(async () => {
      fireEvent.click(screen.getByText('Ask'));
    });

    expect(screen.getByText('Please Confirm')).toBeInTheDocument();
    expect(screen.getByText('Are you sure?')).toBeInTheDocument();
    expect(screen.getByText(CONTINUE)).toBeInTheDocument();
    expect(screen.getByText('Go back')).toBeInTheDocument();

    fireEvent.click(screen.getByText(CONTINUE));

    await waitFor(() => {
      expect(onResult).toHaveBeenCalledWith(true);
    });
    expect(screen.queryByText('Please Confirm')).not.toBeInTheDocument();
  });

  test('clicking Go back resolves false', async () => {
    const onResult = vi.fn();
    render(<ConfirmHost onResult={onResult} />);

    await act(async () => {
      fireEvent.click(screen.getByText('Ask'));
    });

    expect(screen.getByText('Please Confirm')).toBeInTheDocument();

    fireEvent.click(screen.getByText('Go back'));

    await waitFor(() => {
      expect(onResult).toHaveBeenCalledWith(false);
    });
    expect(screen.queryByText('Please Confirm')).not.toBeInTheDocument();
  });

  test('sequential calls to getConfirmation resolve independently', async () => {
    const onResult = vi.fn();
    render(<ConfirmHost onResult={onResult} />);

    await act(async () => {
      fireEvent.click(screen.getByText('Ask'));
    });
    expect(screen.getByText('Please Confirm')).toBeInTheDocument();
    fireEvent.click(screen.getByText(CONTINUE));

    await waitFor(() => {
      expect(onResult).toHaveBeenNthCalledWith(1, true);
    });
    expect(screen.queryByText('Please Confirm')).not.toBeInTheDocument();

    await act(async () => {
      fireEvent.click(screen.getByText('Ask'));
    });
    expect(screen.getByText('Please Confirm')).toBeInTheDocument();
    fireEvent.click(screen.getByText('Go back'));

    await waitFor(() => {
      expect(onResult).toHaveBeenNthCalledWith(2, false);
    });
    expect(screen.queryByText('Please Confirm')).not.toBeInTheDocument();
    expect(onResult).toHaveBeenCalledTimes(2);
  });
});
