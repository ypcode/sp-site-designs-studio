import * as React from 'react';
import { IconButton, Icon, Panel, PanelType } from 'office-ui-fabric-react';
import styles from './ScriptActionAdder.module.scss';
import * as strings from 'SiteDesignsStudioWebPartStrings';
import { ISiteDesignsStudioProps, IServiceConsumerComponentProps } from '../ISiteDesignsStudioProps';
import { escape, assign } from '@microsoft/sp-lodash-subset';

import GenericObjectEditor from '../genericObjectEditor/GenericObjectEditor';
import SiteScriptEditor from '../siteScriptContentEditor/SiteScriptContentEditor';
import { ISiteScriptContent, ISiteScriptAction } from '../../models/ISiteScript';
import {
	ISiteScriptSchemaService,
	SiteScriptSchemaServiceKey
} from '../../services/siteScriptSchema/SiteScriptSchemaService';

export interface IScriptActionAdderState {
	isAdding: boolean;
	availableActions: string[];
}

export interface IScriptActionAdderProps extends IServiceConsumerComponentProps {
	onActionAdded: (s: string) => void;
	parentAction?: ISiteScriptAction;
}

export default class ScriptActionAdder extends React.Component<IScriptActionAdderProps, IScriptActionAdderState> {
	private siteScriptSchemaService: ISiteScriptSchemaService;

	constructor(props: IScriptActionAdderProps) {
		super(props);

		this.props.serviceScope.whenFinished(() => {
			this.siteScriptSchemaService = this.props.serviceScope.consume(SiteScriptSchemaServiceKey);
		});

		this.state = {
			isAdding: false,
			availableActions: []
		};
	}

	public componentWillMount() {
		if (!this.props.parentAction) {
			this.siteScriptSchemaService.getAvailableActionsAsync().then((actions) => {
				this.setState({
					availableActions: actions
				});
			});
		} else {
			this.siteScriptSchemaService.getAvailableSubActionsAsync(this.props.parentAction).then((actions) => {
				this.setState({
					availableActions: actions
				});
			});
		}
	}

	private _addNewAction() {
		this.setState({ isAdding: true });
	}

	private _onPanelDismiss() {
		this.setState({ isAdding: false });
	}

	private _onActionAdded(action: string) {
		this.props.onActionAdded(action);
		this._onPanelDismiss();
	}

	private _translateLabel(value: string): string {
		const key = 'LABEL_' + value;
		return strings[key] || value;
	}

	public render(): React.ReactElement<ISiteDesignsStudioProps> {
		return (
			<div className={styles.scriptActionAdder}>
				<div className={styles.actionAddIcon} onClick={() => this._addNewAction()}>
					<Icon iconName="CircleAdditionSolid" />
				</div>
				<Panel
					type={PanelType.large}
					isOpen={this.state.isAdding}
					headerText="Add a Script Action"
					onDismiss={() => this._onPanelDismiss()}
				>
					<div className="ms-Grid-row">
						{this.state.availableActions.map((a, index) => (
							<div key={index} className="ms-Grid-col ms-sm12 ms-lg6">
								<div className={styles.actionButtonContainer}>
									<div className={styles.actionButton} onClick={() => this._onActionAdded(a)}>
										<div className={styles.actionIcon}>
											<Icon iconName="SetAction" />
										</div>
										<div className={styles.actionButtonLabel}>{this._translateLabel(a)}</div>
									</div>
								</div>
							</div>
						))}
					</div>
				</Panel>
			</div>
		);
	}
}
