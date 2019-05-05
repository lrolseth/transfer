import * as React from "react";
import { IPortalFooterProps, IPortalFooterState } from ".";
import styles from "./PortalFooter.module.scss";
import {
  CommandBar,
  IContextualMenuItem,
  DefaultButton,
  ActionButton,
  Label,
  MessageBar,
  MessageBarType,
  autobind
} from "office-ui-fabric-react";
import * as strings from "ModernFooterApplicationCustomizerStrings";
import { ILinkGroup } from "./ILinkGroup";
import { Links } from "../Links";
import { IPortalFooterEditResult } from "../PortalFooter/IPortalFooterEditResult";

export class PortalFooter extends React.Component<
  IPortalFooterProps,
  IPortalFooterState
> {
  constructor(props: IPortalFooterProps) {
    super(props);

    this.state = {
      expanded: false,
      toggleButtonIconName: "DoubleChevronUp",
      loadingLinks: false,
      links: props.links
    };
  }

  private _handleToggle = (): void => {
    const wasExpanded: boolean = this.state.expanded;

    this.setState({
      expanded: !wasExpanded,
      toggleButtonIconName: wasExpanded
        ? "DoubleChevronUp"
        : "DoubleChevronDown"
    });
  };

  @autobind
  private _handleSupport(): void {
    const supportUrl: string = `mailto:innovators_studio@cargill.com`;
    location.href = supportUrl;
    console.log(supportUrl);
  }

  public render(): React.ReactElement<IPortalFooterProps> {
    return (
      <div className={styles.portalFooter}>
        <Links
          links={this.state.links}
          loadingLinks={this.state.loadingLinks}
          visible={this.state.expanded}
        />
        <div className={styles.main}>
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm3" onClick={this._handleToggle}>
                <div className={styles.logoBackground}>
                  <img
                    alt="Cargill"
                    src="https://cargillonline.sharepoint.com/sites/InnovatorsStudio/SiteAssets/logo.png"
                  />
                </div>
              </div>
              <div className="ms-Grid-col ms-sm4">
                <ActionButton
                  iconProps={{ iconName: "Mail" }}
                  className={styles.supportButton}
                  onClick={this._handleSupport}
                >
                  Contact Innovators Studio
                </ActionButton>
              </div>
              <div className="ms-Grid-col ms-sm4" onClick={this._handleToggle}>
                <Label className={styles.copyright}>
                  Associated Sites, Memberships & Partners
                </Label>
              </div>
              <div className="ms-Grid-col ms-sm1" onClick={this._handleToggle}>
                <div className={styles.toggleControl}>
                  <DefaultButton
                    iconProps={{ iconName: this.state.toggleButtonIconName }}
                    title={
                      this.state.expanded
                        ? strings.ToggleButtonClose
                        : strings.ToggleButtonOpen
                    }
                    className={styles.toggleButton}
                    onClick={this._handleToggle}
                  />
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
