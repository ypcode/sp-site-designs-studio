import * as React from 'react';
import { Dropdown, TextField, Toggle, Link, IconButton } from 'office-ui-fabric-react';
import styles from './ScriptActionEditor.module.scss';
import { escape, assign } from '@microsoft/sp-lodash-subset';
import * as strings from 'SiteDesignsStudioWebPartStrings';
import GenericObjectEditor from '../genericObjectEditor/GenericObjectEditor';

import { ISiteScriptAction } from '../../models/ISiteScript';
import ScriptActionAdder from '../scriptActionAdder/ScriptActionAdder';
import { IServiceConsumerComponentProps } from '../ISiteDesignsStudioProps';
import {
	ISiteScriptSchemaService,
	SiteScriptSchemaServiceKey
} from '../../services/siteScriptSchema/SiteScriptSchemaService';
import { ISiteDesignsService, SiteDesignsServiceKey } from '../../services/siteDesigns/SiteDesignsService';

export interface IScriptActionEditorState {
	expandedSubactionIndices: number[];
}

export interface IScriptActionEditorProps extends IServiceConsumerComponentProps {
	action: ISiteScriptAction;
	schema: any;
	isExpanded: boolean;
	onExpandToggle?: (isExpanded: boolean) => void;
	onActionChanged: (action: ISiteScriptAction) => void;
	onRemove: () => void;
	getActionName: (action: ISiteScriptAction) => string;
}

export default class ScriptActionEditor extends React.Component<IScriptActionEditorProps, IScriptActionEditorState> {
	private siteScriptSchemaService: ISiteScriptSchemaService;
	private siteDesignsService: ISiteDesignsService;

	constructor(props: IScriptActionEditorProps) {
		super(props);

		this.props.serviceScope.whenFinished(() => {
			this.siteScriptSchemaService = this.props.serviceScope.consume(SiteScriptSchemaServiceKey);
			this.siteDesignsService = this.props.serviceScope.consume(SiteDesignsServiceKey);
		});

		this.state = {
			expandedSubactionIndices: []
		};
	}

	// public componentWillReceiveProps(nextProps: IScriptActionEditorProps) {
	// 	this._setAllSubactionsExpanded(nextProps.isExpanded);
	// }

	private _toggleIsExpanded() {
		if (this.props.onExpandToggle) {
			this.props.onExpandToggle(!this.props.isExpanded);
		}
	}

	private _translateLabel(value: string): string {
		const key = 'LABEL_' + value;
		return strings[key] || value;
	}

	private _onSubActionChanged(parentAction: ISiteScriptAction, subAction: ISiteScriptAction) {
		let subactions = parentAction['subactions'] as ISiteScriptAction[];
		parentAction['subactions'] = [].concat(subactions);
		this.props.onActionChanged(parentAction);
	}

	private _setSubActionExpanded(actionIndex: number, isExpanded: boolean) {
		let { expandedSubactionIndices } = this.state;
		let woCurrentIndex = expandedSubactionIndices.filter((i) => i != actionIndex);
		this.setState({
			expandedSubactionIndices: isExpanded ? woCurrentIndex.concat(actionIndex) : woCurrentIndex
		});
	}

	private _setAllSubactionsExpanded(isExpanded: boolean) {
		let { action } = this.props;
		let { expandedSubactionIndices } = this.state;
		if (action.subactions) {
			this.setState({
				expandedSubactionIndices: isExpanded ? action.subactions.map((item, index) => index) : []
			});
		}
	}

	private _setSingleSubactionExpanded(actionIndex: number) {
		this.setState({
			expandedSubactionIndices: [ actionIndex ]
		});
	}

	private _isSubactionExpanded(index: number): boolean {
		return this.state.expandedSubactionIndices.indexOf(index) > -1;
	}

	public render(): React.ReactElement<IScriptActionEditorProps> {
		let { isExpanded, action, serviceScope, schema, onActionChanged } = this.props;
		let expandCollapseIcon = isExpanded ? 'CollapseContentSingle' : 'ExploreContentSingle';

		const subactionsRenderer = (subactions: ISiteScriptAction[]) => (
			<div className={styles.subactions}>
				<h3>{this._translateLabel('subactions')}</h3>
				<div className={styles.subactionsWorkspace}>
					{subactions.map((subaction, index) => (
						<div className={styles.subactionItem}>
							<ScriptActionEditor
								key={`SUBACTION_${index}`}
								serviceScope={this.props.serviceScope}
								isExpanded={this._isSubactionExpanded(index)}
								onExpandToggle={(expanded) => this._setSubActionExpanded(index, expanded)}
								action={subaction}
								getActionName={(s) => s.verb}
								schema={this.siteScriptSchemaService.getSubActionSchema(action, subaction)}
								onRemove={() => this._removeScriptSubAction(action, index)}
								onActionChanged={(a) => this._onSubActionUpdated(action, index, a)}
							/>
						</div>
					))}
					<div>
						<ScriptActionAdder
							parentAction={action}
							serviceScope={serviceScope}
							onActionAdded={(a) => this._addScriptSubAction(action, a)}
						/>
					</div>
				</div>
			</div>
		);

		return (
			<div className={styles.scriptActionEditor}>
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm10">
						<h2 className={styles.title}>
							{this._translateLabel(this.props.getActionName(this.props.action))}
						</h2>
					</div>
					<div className="ms-Grid-col ms-sm2 close">
						<IconButton
							iconProps={{ iconName: expandCollapseIcon }}
							onClick={() => this._toggleIsExpanded()}
						/>
						<IconButton iconProps={{ iconName: 'ChromeClose' }} onClick={() => this.props.onRemove()} />
					</div>
				</div>
				{isExpanded && (
					<div className="ms-Grid-row">
						<div className="ms-Grid-col ms-sm12">
							<GenericObjectEditor
								customRenderers={{ subactions: subactionsRenderer }}
								defaultValues={{ subactions: [] }}
								object={action}
								schema={schema}
								ignoredProperties={[ 'verb' ]}
								onObjectChanged={onActionChanged.bind(this)}
							/>
						</div>
					</div>
				)}
			</div>
		);
	}

	private _addScriptSubAction(parentAction: ISiteScriptAction, verb: string) {
		let newSubAction: ISiteScriptAction = {
			verb: verb
		};

		let newSubActions = [].concat(parentAction.subactions, newSubAction);
		let updatedAction = assign({}, parentAction);
		updatedAction.subactions = newSubActions;
		this._setSingleSubactionExpanded(newSubActions.length - 1);
		this.props.onActionChanged(updatedAction);
	}

	private _removeScriptSubAction(parentAction: ISiteScriptAction, subActionKey: number) {
		let newSubActions = parentAction.subactions.filter((item, index) => index != subActionKey);
		let updatedAction = assign({}, parentAction);
		updatedAction.subactions = newSubActions;
		this.props.onActionChanged(updatedAction);
	}

	private _onSubActionUpdated(
		parentAction: ISiteScriptAction,
		subActionKey: number,
		updatedSubAction: ISiteScriptAction
	) {
		let subAction = assign({}, parentAction.subactions[subActionKey], updatedSubAction);

		let updatedParentAction = assign({}, parentAction);
		updatedParentAction.subactions = parentAction.subactions.map(
			(sa, ndx) => (ndx == subActionKey ? subAction : sa)
		);
		this.props.onActionChanged(updatedParentAction);
	}
}
