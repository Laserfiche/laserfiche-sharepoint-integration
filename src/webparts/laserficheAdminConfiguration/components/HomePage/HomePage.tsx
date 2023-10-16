import * as React from 'react';
require('./../../../../Assets/CSS/commonStyles.css');
require('../../../../Assets/CSS/bootstrap.min.css');

const YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART =
  'You must be a currently licensed Laserfiche Cloud user to use this web part.';
const FOR_MORE_INFO_VISIT = 'For more information visit';
const USE_LASERFICHE_ADMIN_PAGE_TO_EDIT_SP_LF_CONFIG_SIGN_IN_AND_SELECT_TASK_TO_COMPLETE =
  'Use the Laserfiche Administration page to edit your SharePoint and Laserfiche configuration. Sign in and select the task you want to perform from the menu at the top of this section.';
const PROFILES_GOVERN_HOW_SP_SAVED_TO_LF_CAN_CREATE_MULTIPLE_PROFILES_FOR_DIFFERENT_SP_CONTENT_TYPES =
  'Profiles govern how documents in SharePoint will be saved to Laserfiche. You can create multiple profiles for different SharePoint content types. For example, if you want applications stored differently than invoices, create separate profiles for each content type.';
const IN_THIS_TAB_CAN_MAP_SPECIFIC_SP_CONTENT_TYPE_TO_LF_PROFILE_PROFILE_THEN_USED_TO_SAVE_SP_DOCS_WITH_THAT_CONTENT_TYPE =
  'In this tab, you can map a specific SharePoint content type to a corresponding Laserfiche profile. This profile will then be used when saving all documents of the specified SharePoint content type.';

export default function HomePage(): JSX.Element {
  return (
    <div className='p-3'>
      <main className='bg-white shadow-sm'>
        <p className='adminContent'>
          {`${YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART} ${FOR_MORE_INFO_VISIT} `}
          <a href='https://www.laserfiche.com/products/pricing'>
            laserfiche.com
          </a>
          {`.`}
        </p>
        <p className='adminContent'>
          {
            USE_LASERFICHE_ADMIN_PAGE_TO_EDIT_SP_LF_CONFIG_SIGN_IN_AND_SELECT_TASK_TO_COMPLETE
          }
        </p>
        <p className='adminContent'>
          For more information, see the{' '}
          <a
            href='https://laserfiche.github.io/laserfiche-sharepoint-integration/'
            target='_blank'
            rel='noreferrer'
            style={{ color: '#0079d6' }}
          >
            help documentation.
          </a>
        </p>
        <div className='adminContent'>
          <p>
            <strong>Profiles</strong>
          </p>
          <p style={{ marginLeft: '38px' }}>
            <span>
              {
                PROFILES_GOVERN_HOW_SP_SAVED_TO_LF_CAN_CREATE_MULTIPLE_PROFILES_FOR_DIFFERENT_SP_CONTENT_TYPES
              }
            </span>
          </p>
          <p>
            <strong>Profile Mapping</strong>
          </p>
          <p style={{ marginLeft: '38px' }}>
            <span>
              {
                IN_THIS_TAB_CAN_MAP_SPECIFIC_SP_CONTENT_TYPE_TO_LF_PROFILE_PROFILE_THEN_USED_TO_SAVE_SP_DOCS_WITH_THAT_CONTENT_TYPE
              }
            </span>
          </p>
        </div>
      </main>
    </div>
  );
}
