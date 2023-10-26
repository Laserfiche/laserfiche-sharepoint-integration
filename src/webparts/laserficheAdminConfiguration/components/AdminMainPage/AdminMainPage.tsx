// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import { NavLink } from 'react-router-dom';
import { useEffect } from 'react';
import { IAdminPageProps } from './IAdminPageProps';
import { CreateConfigurations } from '../../../../Utils/CreateConfigurations';
require('../../../../Assets/CSS/bootstrap.min.css');
require('./../../../../Assets/CSS/commonStyles.css');
import styles from './../LaserficheAdminConfiguration.module.scss';

declare global {
  // eslint-disable-next-line
  namespace JSX {
    interface IntrinsicElements {
      // eslint-disable-next-line
      ['lf-login']: any;
    }
  }
}

export default function AdminMainPage(props: IAdminPageProps): JSX.Element {
  useEffect(() => {
    CreateConfigurations.ensureAdminConfigListCreatedAsync(props.context).catch(
      (err: Error) => {
        console.warn(
          `Error: ${err.message}`
        );
      }
    );
  }, []);

  const linkData: LinkInfo[] = [
    { route: '/HomePage', name: 'About' },
    { route: '/ManageConfigurationsPage', name: 'Profiles' },
    { route: '/ManageMappingsPage', name: 'Profile Mapping' },
  ];

  return (
    <div style={{ borderBottom: '3px solid #CE7A14' }}>
      <div>
        <span className={styles.profileTitle}>Profile Editor</span>
        {props.loggedIn && <Links linkData={linkData} />}
      </div>
    </div>
  );
}

interface LinkInfo {
  route: string;
  name: string;
}

function Links(props: { linkData: LinkInfo[] }): JSX.Element {
  const linkEls = props.linkData.map((link: LinkInfo) => (
    <span key={link.name}>
      <NavLink
        to={link.route}
        activeStyle={{ fontWeight: 'bold', textDecoration: 'underline' }}
        style={{
          marginRight: '25px',
          fontWeight: '500',
          fontSize: '15px',
          color: '#0079d6',
        }}
      >
        {link.name}
      </NavLink>
    </span>
  ));
  return <div>{linkEls}</div>;
}
