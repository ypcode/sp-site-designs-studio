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
	allSubactionsExpanded: boolean;
	onExpandToggle?: (isExpanded: boolean) => void;
	onActionChanged?: (action: ISiteScriptAction) => void;
	onRemove?: () => void;
	canMoveUp?: boolean;
	onMoveUp?: () => void;
	canMoveDown?: boolean;
	onMoveDown?: () => void;
	getActionName?: (action: ISiteScriptAction) => string;
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

	public componentWillReceiveProps(nextProps: IScriptActionEditorProps) {
		this._setAllSubactionsExpanded(nextProps.allSubactionsExpanded);
	}

	private _toggleIsExpanded() {
		if (this.props.onExpandToggle) {
			this.props.onExpandToggle(!this.props.isExpanded);
		}
	}

	// TODO Reuse the current private method from schema service
	private _getVerbFromActionSchema(actionDefinition: any): string {
		if (
			!actionDefinition.properties ||
			!actionDefinition.properties.verb ||
			!actionDefinition.properties.verb.enum ||
			!actionDefinition.properties.verb.enum.length
		) {
			throw new Error('Invalid Action schema');
		}

		return actionDefinition.properties.verb.enum[0];
	}

	private _getCurrentActionName(): string {
		let { schema, action, getActionName } = this.props;
		if (getActionName) {
			return getActionName(action);
		} else {
			return this._getVerbFromActionSchema(schema);
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
		let {
			isExpanded,
			action,
			serviceScope,
			schema,
			onActionChanged,
			allSubactionsExpanded,
			canMoveDown,
			canMoveUp
		} = this.props;
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
								allSubactionsExpanded={allSubactionsExpanded}
								onExpandToggle={(expanded) => this._setSubActionExpanded(index, expanded)}
								action={subaction}
								getActionName={this.props.getActionName}
								schema={this.siteScriptSchemaService.getSubActionSchema(action, subaction)}
								onRemove={() => this._removeScriptSubAction(action, index)}
								onActionChanged={(a) => this._onSubActionUpdated(action, index, a)}
								canMoveDown={index < subactions.length - 1}
								onMoveDown={() => this._moveSubactionDown(action, index)}
								canMoveUp={index > 0}
								onMoveUp={() => this._moveSubactionUp(action, index)}
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
					<div className="ms-Grid-col ms-sm8">
						<h2 className={styles.title}>
							{this._translateLabel(this._getCurrentActionName())}
						</h2>
					</div>
					<div className="ms-Grid-col ms-sm4">
						<div className={styles.commandButtonsContainer}>
							<div className={styles.commandButtons}>
								<IconButton
									iconProps={{ iconName: 'Up' }}
									onClick={() => this._onMoveUp()}
									disabled={!canMoveUp}
								/>
								<IconButton
									iconProps={{ iconName: 'Down' }}
									onClick={() => this._onMoveDown()}
									disabled={!canMoveDown}
								/>
								<IconButton
									iconProps={{ iconName: expandCollapseIcon }}
									onClick={() => this._toggleIsExpanded()}
								/>
								<IconButton
									iconProps={{ iconName: 'ChromeClose' }}
									onClick={() => this.props.onRemove()}
								/>
							</div>
						</div>
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

	private _onMoveUp() {
		if (this.props.onMoveUp) {
			this.props.onMoveUp();
		}
	}

	private _onMoveDown() {
		if (this.props.onMoveDown) {
			this.props.onMoveDown();
		}
	}

	private _moveSubactionUp(parentAction: ISiteScriptAction, index: number) {
		this._swapSubActions(parentAction, index, index - 1);
	}

	private _moveSubactionDown(parentAction: ISiteScriptAction, index: number) {
		this._swapSubActions(parentAction, index, index + 1);
	}

	private _swapSubActions(parentAction: ISiteScriptAction, oldIndex: number, newIndex: number) {
		if (newIndex < 0 || newIndex > parentAction.subactions.length - 1) {
			return;
		}

		let newSubActions = [].concat(parentAction.subactions);

		let old = newSubActions[oldIndex];
		newSubActions[oldIndex] = newSubActions[newIndex];
		newSubActions[newIndex] = old;

		let updatedAction = assign({}, parentAction);
		updatedAction.subactions = newSubActions;
		this.props.onActionChanged(updatedAction);
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
