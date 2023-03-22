import * as React from 'react';
import { ILaserficheAdminConfigurationProps } from './ILaserficheAdminConfigurationProps';
import { HashRouter, Route, Switch } from 'react-router-dom';
import { Stack, StackItem } from 'office-ui-fabric-react';
import AdminMainPage from '../components/AdminMainPage/AdminMainPage';
import HomePage from './HomePage/HomePage';
import ManageConfigurationsPage from './ManageConfigurationsPage/ManageConfigurationsPage';
import ManageMappingsPage from './ManageMappingsPage/ManageMappingsPage';
import EditManageConfiguration from './EditManageConfiguration/EditManageConfiguration';
import AddNewManageConfiguration from './AddNewManageConfiguration/AddNewManageConfiguration';


export default class LaserficheAdminConfiguration extends React.Component<ILaserficheAdminConfigurationProps, {}> {
  constructor(props: ILaserficheAdminConfigurationProps) {
    super(props);
  }
  public render(): React.ReactElement<ILaserficheAdminConfigurationProps> {
    return (
      <HashRouter>
        <Stack>
          <AdminMainPage context={this.props.context} webPartTitle={this.props.webPartTitle} laserficheRedirectPage={this.props.laserficheRedirectPage} devMode={this.props.devMode} ></AdminMainPage>
          <StackItem>
            <Switch>
              <Route exact={true} component={() => <HomePage />} path="/HomePage">
              </Route>
              <Route exact={true} component={() => <HomePage />} path="/">
              </Route>
              <Route exact={true} component={() => <ManageConfigurationsPage context={this.props.context} />} path="/ManageConfigurationsPage">
              </Route>
              <Route exact={true} component={() => <ManageMappingsPage context={this.props.context} />} path="/ManageMappingsPage" >
              </Route>
              <Route exact={true} component={() => <AddNewManageConfiguration context={this.props.context} laserficheRedirectPage={this.props.laserficheRedirectPage} devMode={this.props.devMode}  />} path="/AddNewManageConfiguration" >
              </Route>
              <Route exact={true} render={(props) => <EditManageConfiguration {...props} context={this.props.context} laserficheRedirectPage={this.props.laserficheRedirectPage} devMode={this.props.devMode} />} path="/EditManageConfiguration/:name" >
              </Route>
            </Switch>
          </StackItem>
        </Stack>
      </HashRouter>
    );
  }
}
