// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

vi.mock('../../../../Utils/CreateConfigurations');

import type { Mock } from 'vitest';
import * as React from 'react';
import { render, screen, waitFor } from '@testing-library/react';
import { HashRouter } from 'react-router-dom';
import AdminMainPage from './AdminMainPage';
import { IAdminPageProps } from './IAdminPageProps';
import { CreateConfigurations } from '../../../../Utils/CreateConfigurations';

describe('AdminMainPage', () => {
  const buildProps = (loggedIn: boolean): IAdminPageProps =>
    ({
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      context: {} as any,
      loggedIn,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      repoClient: {} as any,
    }) as IAdminPageProps;

  beforeEach(() => {
    vi.clearAllMocks();
    (CreateConfigurations.ensureAdminConfigListCreatedAsync as Mock).mockResolvedValue(undefined);
  });

  test('renders "Profile Editor"', () => {
    // Act
    render(
      <HashRouter>
        <AdminMainPage {...buildProps(true)} />
      </HashRouter>
    );

    // Assert
    expect(screen.getByText('Profile Editor')).toBeInTheDocument();
  });

  test('does not render the links when loggedIn is false', () => {
    // Act
    render(
      <HashRouter>
        <AdminMainPage {...buildProps(false)} />
      </HashRouter>
    );

    // Assert
    expect(screen.queryByText('About')).not.toBeInTheDocument();
    expect(screen.queryByText('Profiles')).not.toBeInTheDocument();
    expect(screen.queryByText('Profile Mapping')).not.toBeInTheDocument();
  });

  test('renders all links with the expected hrefs when loggedIn is true', () => {
    // Act
    render(
      <HashRouter>
        <AdminMainPage {...buildProps(true)} />
      </HashRouter>
    );

    // Assert
    const about = screen.getByText('About');
    const profiles = screen.getByText('Profiles');
    const profileMapping = screen.getByText('Profile Mapping');

    expect(about.tagName).toBe('A');
    expect(profiles.tagName).toBe('A');
    expect(profileMapping.tagName).toBe('A');

    expect((about as HTMLAnchorElement).href).toMatch(/#\/HomePage$/);
    expect((profiles as HTMLAnchorElement).href).toMatch(/#\/ManageConfigurationsPage$/);
    expect((profileMapping as HTMLAnchorElement).href).toMatch(/#\/ManageMappingsPage$/);
  });

  test('calls ensureAdminConfigListCreatedAsync exactly once with props.context on mount', () => {
    // Arrange
    const props = buildProps(true);

    // Act
    render(
      <HashRouter>
        <AdminMainPage {...props} />
      </HashRouter>
    );

    // Assert
    expect(CreateConfigurations.ensureAdminConfigListCreatedAsync).toHaveBeenCalledTimes(1);
    expect(CreateConfigurations.ensureAdminConfigListCreatedAsync).toHaveBeenCalledWith(
      props.context
    );
  });

  test('calls ensureAdminConfigListCreatedAsync once even when loggedIn is false', () => {
    // Arrange
    const props = buildProps(false);

    // Act
    render(
      <HashRouter>
        <AdminMainPage {...props} />
      </HashRouter>
    );

    // Assert
    expect(CreateConfigurations.ensureAdminConfigListCreatedAsync).toHaveBeenCalledTimes(1);
    expect(CreateConfigurations.ensureAdminConfigListCreatedAsync).toHaveBeenCalledWith(
      props.context
    );
  });

  test('logs a console.warn and still renders when ensureAdminConfigListCreatedAsync rejects', async () => {
    // Arrange
    (CreateConfigurations.ensureAdminConfigListCreatedAsync as Mock).mockRejectedValue(
      new Error('boom')
    );
    const consoleWarnSpy = vi.spyOn(console, 'warn').mockImplementation(() => undefined);

    // Act
    render(
      <HashRouter>
        <AdminMainPage {...buildProps(true)} />
      </HashRouter>
    );

    // Assert
    await waitFor(() => {
      expect(consoleWarnSpy).toHaveBeenCalledWith(expect.stringContaining('Error: "boom"'));
    });
    expect(screen.getByText('Profile Editor')).toBeInTheDocument();
  });
});
