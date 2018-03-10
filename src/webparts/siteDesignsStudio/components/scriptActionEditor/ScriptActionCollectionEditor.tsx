import * as React from 'react';
import { SortableContainer, SortableHandle, SortableElement } from 'react-sortable-hoc';
import { Dropdown, TextField, Toggle, Link, IconButton } from 'office-ui-fabric-react';
import styles from './ScriptActionEditor.module.scss';
import { escape, assign } from '@microsoft/sp-lodash-subset';
import * as strings from 'SiteDesignsStudioWebPartStrings';

import { ISiteScriptAction } from '../../models/ISiteScript';
import ScriptActionAdder from '../scriptActionAdder/ScriptActionAdder';
import { IServiceConsumerComponentProps } from '../ISiteDesignsStudioProps';
import {
	ISiteScriptSchemaService,
	SiteScriptSchemaServiceKey
} from '../../services/siteScriptSchema/SiteScriptSchemaService';
import { ISiteDesignsService, SiteDesignsServiceKey } from '../../services/siteDesigns/SiteDesignsService';
import ScriptActionEditor from './ScriptActionEditor';

interface ISortStartEventArgs {
	node: any;
	index: number;
	collection: any[];
}

interface ISOrtEndEventArgs {
	oldIndex: number;
	newIndex: number;
	collection: any[];
}

export interface IScriptActionCollectionEditorState {}

export interface IScriptActionCollectionEditorProps extends IServiceConsumerComponentProps {
	actions: ISiteScriptAction[];
	onActionChanged?: (actionIndex: number, action: ISiteScriptAction) => void;
	onActionRemoved?: (actionIndex: number) => void;
	onActionMoved?: (oldActionIndex: number, newActionIndex: number) => void;
	expandedIndices: number[];
	onExpandChanged?: (expandedIndices: number[], parentAction?: ISiteScriptAction) => void;
	getActionSchema?: (action: ISiteScriptAction) => any;
}

export default class ScriptActionCollectionEditor extends React.Component<
	IScriptActionCollectionEditorProps,
	IScriptActionCollectionEditorState
> {
	private siteScriptSchemaService: ISiteScriptSchemaService;
	private siteDesignsService: ISiteDesignsService;

	constructor(props: IScriptActionCollectionEditorProps) {
		super(props);

		this.props.serviceScope.whenFinished(() => {
			this.siteScriptSchemaService = this.props.serviceScope.consume(SiteScriptSchemaServiceKey);
			this.siteDesignsService = this.props.serviceScope.consume(SiteDesignsServiceKey);
		});
	}

	private _translateLabel(value: string): string {
		const key = 'LABEL_' + value;
		return strings[key] || value;
	}

	public render(): React.ReactElement<IScriptActionCollectionEditorProps> {
		let { actions, serviceScope, onActionChanged } = this.props;
		console.log('ACTIONS= ', actions);

		const SortableListContainer = SortableContainer(({ items }) => {
			return <div>{items.map((value, index) => this._renderActionEditorWithCommands(value, index))}</div>;
		});

		return (
			<SortableListContainer
				items={actions}
				onSortStart={(args) => this._onSortStart(args)}
        onSortEnd={(args) => this._onSortEnd(args)}
        lockToContainerEdges={true}
				useDragHandle={true}
			/>
		);
	}

  // private sortedItemIsExpanded: boolean;
  // private isSorting: boolean = null;
	private _onSortStart(args: ISortStartEventArgs) {
    // this.sortedItemIsExpanded = this._isExpanded(args.index);
    // this.isSorting = true;
	}

	private _onSortEnd(args: ISOrtEndEventArgs) {
    let wasPreviousExpanded = this._isExpanded(args.oldIndex);
		this._moveAction(args.oldIndex, args.newIndex);
    // Set the initial collapse status of the item
    // this.isSorting = false;
    // this._setExpanded(args.newIndex, this.sortedItemIsExpanded);
    // this._setExpanded(args.oldIndex, wasPreviousExpanded)
    // this.sortedItemIsExpanded = null;
	}

	private _setExpanded(actionIndex: number, expanded: boolean) {
		let { expandedIndices } = this.props;
		let expandedWoCurrent = expandedIndices.filter((i) => i != actionIndex);
		expandedIndices = expanded ? expandedWoCurrent.concat(actionIndex) : expandedWoCurrent;

		if (this.props.onExpandChanged) {
			this.props.onExpandChanged(expandedIndices);
		}
	}

	private _renderActionEditorWithCommands(action: ISiteScriptAction, actionIndex: number) {
		let { expandedIndices, getActionSchema } = this.props;
		let actionSchema = getActionSchema(action);

		const DragHandle = SortableHandle(() => (
			<h2 className={styles.title}>{this._getActionNameFromActionSchema(actionSchema)}</h2>
		));

		let isExpanded = this._isExpanded(actionIndex);
		let expandCollapseIcon = isExpanded ? 'CollapseContentSingle' : 'ExploreContentSingle';
		const SortableItem = SortableElement(({ value }) => (
			<div>
				<div className={styles.scriptActionEditor}>
					<div className="ms-Grid-row">
						<div className="ms-Grid-col ms-sm8">
							<DragHandle />
						</div>
						<div className="ms-Grid-col ms-sm4">
							<div className={styles.commandButtonsContainer}>
								<div className={styles.commandButtons}>
									<IconButton
										iconProps={{ iconName: expandCollapseIcon }}
										onClick={() => this._toggleExpanded(actionIndex)}
									/>
									<IconButton
										iconProps={{ iconName: 'ChromeClose' }}
										onClick={() => this._removeAction(actionIndex)}
									/>
								</div>
							</div>
						</div>
					</div>
					{isExpanded && (
						<ScriptActionEditor
							allSubactionsExpanded={true}
							serviceScope={this.props.serviceScope}
							action={value}
							schema={actionSchema}
							expandedSubActions={[]}
							onActionChanged={(updated) => this._onActionUpdated(actionIndex, updated)}
						/>
					)}
				</div>
			</div>
		));

		return <SortableItem key={`item-${actionIndex}`} index={actionIndex} value={action} />;
	}

	private _getActionNameFromActionSchema(actionSchema: any): string {
		return this._translateLabel(this._getVerbFromActionSchema(actionSchema));
	}

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

	private _isExpanded(actionIndex: number): boolean {
		let { expandedIndices } = this.props;
		return expandedIndices.indexOf(actionIndex) > -1;
	}
	private _toggleExpanded(actionIndex: number) {
		let isExpanded = this._isExpanded(actionIndex);
		this._setExpanded(actionIndex, !isExpanded);
	}

	private _removeAction(actionIndex: number) {
		if (this.props.onActionRemoved) {
			this.props.onActionRemoved(actionIndex);
		}
	}

	private _onActionUpdated(actionIndex: number, action: ISiteScriptAction) {
		if (this.props.onActionChanged) {
			this.props.onActionChanged(actionIndex, action);
		}
	}
	private _moveAction(oldIndex: number, newIndex: number) {
		if (this.props.onActionMoved) {
			this.props.onActionMoved(oldIndex, newIndex);
		}
	}
}
