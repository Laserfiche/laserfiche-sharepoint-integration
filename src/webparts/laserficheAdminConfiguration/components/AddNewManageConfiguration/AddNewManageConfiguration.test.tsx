// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import '@testing-library/jest-dom';
import { render, screen } from '@testing-library/react';
import AddNewManageConfiguration from './AddNewManageConfiguration';
import mockWebPartContext from '../../../../__mocks__/@microsoft/sp-webpart-base';
import { RepositoryClientExInternal } from '../../../../repository-client/repository-client';
import { BrowserRouter } from 'react-router-dom';

const MANAGE_CONFIGURATIONS_PAGE_TITLE = "Mappings from SharePoint Column to Laserfiche Field Values";
test('renders title of page', () => {
  render(
    <BrowserRouter>
      <AddNewManageConfiguration
        context={mockWebPartContext}
        repoClient={new RepositoryClientExInternal().repoClient}
        loggedIn={true}
      />
    </BrowserRouter>
  );
  const linkElement = screen.getByText(
    MANAGE_CONFIGURATIONS_PAGE_TITLE
  );
  expect(linkElement).toBeInTheDocument();
});
