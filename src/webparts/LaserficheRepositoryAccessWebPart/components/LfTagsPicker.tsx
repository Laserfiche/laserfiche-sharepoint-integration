// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { NgElement, WithProperties } from '@angular/elements';
import { LfRepoTagsService } from '@laserfiche/lf-ui-components-services';
import { LfTagDefinition, LfTagsComponent } from '@laserfiche/types-lf-ui-components';
import * as React from 'react';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import styles from './LaserficheRepositoryAccess.module.scss';

export function LfTagsPicker(props: {
  repoClient: IRepositoryApiClientExInternal;
  onSelectedTagNamesChange: (tagNames: string[]) => void;
}): JSX.Element {
  const host = React.useRef<HTMLDivElement>();

  React.useEffect(() => {
    const onSelectedTagsChanged: EventListener = (event: Event) => {
      props.onSelectedTagNamesChange(
        (event as CustomEvent<LfTagDefinition[]>).detail.map((tag) => tag.displayName)
      );
    };

    // <lf-tags> asks its tagsService for tags only once, as it's attached to
    // the page, so the element is built here with the service already set --
    // React's JSX would attach it before any ref or effect could set it.
    const tags = document.createElement('lf-tags') as NgElement & WithProperties<LfTagsComponent>;
    tags.tagsService = new LfRepoTagsService(props.repoClient);
    tags.addEventListener('selectedTagsChanged', onSelectedTagsChanged);
    host.current.appendChild(tags);

    return () => {
      tags.removeEventListener('selectedTagsChanged', onSelectedTagsChanged);
      tags.remove();
    };
  }, [props.repoClient]);

  return <div className={styles.lfComponentContainer} ref={host} />;
}
