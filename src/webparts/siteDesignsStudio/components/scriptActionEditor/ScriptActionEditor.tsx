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
import ScriptActionCollectionEditor from './ScriptActionCollectionEditor';

export interface IScriptActionEditorState {

}

export interface IScriptActionEditorProps extends IServiceConsumerComponentProps {
	action: ISiteScriptAction;
	schema: any;
  allSubactionsExpanded: boolean;
  expandedSubActions: number[];
  onSubActionsExpandChanged?: (expanded: number[]) => void;
	onActionChanged?: (action: ISiteScriptAction) => void;
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
		};
	}

	public componentWillReceiveProps(nextProps: IScriptActionEditorProps) {
		// this._setAllSubactionsExpanded(nextProps.allSubactionsExpanded);
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
		let { schema } = this.props;
		return this._getVerbFromActionSchema(schema);
	}

	private _translateLabel(value: string): string {
		const key = 'LABEL_' + value;
		return strings[key] || value;
	}

	private _onSubActionAdded(parentAction: ISiteScriptAction, subAction: ISiteScriptAction) {
		let subactions = parentAction['subactions'] as ISiteScriptAction[];
		parentAction['subactions'] = [].concat(subactions);
		this.props.onActionChanged(parentAction);
  }

  private _getExpandedSubActions() : number[] {
    let {allSubactionsExpanded, action, expandedSubActions} = this.props;
    if (action.subactions && action.subactions.length  == 0) {
      return [];
    }

    if (this.props.allSubactionsExpanded) {
      return action.subactions.map((a, i) => i);
    }

    return expandedSubActions;
  }

	public render(): React.ReactElement<IScriptActionEditorProps> {
		let { action, serviceScope, schema, onActionChanged, expandedSubActions } = this.props;

		const subactionsRenderer = (subactions: ISiteScriptAction[]) => (
			<div className={styles.subactions}>
				<h3>{this._translateLabel('subactions')}</h3>
				<div className={styles.subactionsWorkspace}>
					<div>
						<ScriptActionCollectionEditor
							serviceScope={this.props.serviceScope}
							actions={subactions}
							onActionRemoved={(subActionIndex) => this._removeScriptSubAction(action, subActionIndex)}
							onActionMoved={(oldIndex, newIndex) => this._moveSubAction(action, oldIndex, newIndex)}
							onActionChanged={(subActionIndex, subAction) =>
								this._onSubActionUpdated(action, subActionIndex, subAction)}
							getActionSchema={(subAction) =>
								this.siteScriptSchemaService.getSubActionSchema(action, subAction)}
							expandedIndices={this._getExpandedSubActions()}
							onExpandChanged={(expanded) => this._onSubActionsExpandChanged(expanded)}
						/>
					</div>
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
			<div className="ms-Grid-row">
				<div className="ms-Grid-col ms-sm12">
					<GenericObjectEditor
						customRenderers={{ subactions: subactionsRenderer }}
						defaultValues={{ subactions: [] }}
						object={action}
						schema={schema}
						ignoredProperties={[ 'verb' ]}
            onObjectChanged={onActionChanged.bind(this)}
            updateOnBlur={true}
					/>
				</div>
			</div>
		);
	}

	private _moveSubAction(parentAction: ISiteScriptAction, oldIndex: number, newIndex: number) {
		if (newIndex < 0 || newIndex > parentAction.subactions.length - 1) {
			return;
		}

		let newSubActions = [].concat(parentAction.subactions);

    let actionToMove = newSubActions[oldIndex];
    newSubActions.splice(oldIndex, 1);
    newSubActions.splice(newIndex, 0, actionToMove);

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
		// this._setSingleSubactionExpanded(newSubActions.length - 1);
		this.props.onActionChanged(updatedAction);
	}

	private _removeScriptSubAction(parentAction: ISiteScriptAction, subActionKey: number) {
		let newSubActions = parentAction.subactions.filter((item, index) => index != subActionKey);
		let updatedAction = assign({}, parentAction);
		updatedAction.subactions = newSubActions;
		this.props.onActionChanged(updatedAction);
	}

	private _onSubActionsExpandChanged(expanded: number[]) {
		if (this.props.onSubActionsExpandChanged) {
      this.props.onSubActionsExpandChanged(expanded);
    }
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
