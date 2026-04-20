import { AgentToolExecutor } from '../../tools/executor'
import {
  createWorkspaceToolPack,
  type WorkspaceToolPackDeps,
} from '../../tools/packs/workspace'
import { createWebToolPack, type WebToolPackDeps } from '../../tools/packs/web'
import { createWordToolPack, type WordToolPackDeps } from '../../tools/packs/word'
import { createExcelToolPack, type ExcelToolPackDeps } from '../../tools/packs/excel'
import { createPptToolPack, type PptToolPackDeps } from '../../tools/packs/ppt'
import { createMemoryToolPack, type MemoryToolPackDeps } from '../../tools/packs/memory'
import { createKnowledgeToolPack, type KnowledgeToolPackDeps } from '../../tools/packs/knowledge'
import { AgentToolRegistry } from '../../tools/registry'
import { AgentToolScheduler } from '../../tools/scheduler'

export interface PhaseOneToolExecutorDeps {
  word: WordToolPackDeps
  excel: ExcelToolPackDeps
  ppt: PptToolPackDeps
  workspace: WorkspaceToolPackDeps
  web: WebToolPackDeps
  memory: MemoryToolPackDeps
  knowledge: KnowledgeToolPackDeps
}

export function createPhaseOneToolExecutor(
  deps: PhaseOneToolExecutorDeps,
): AgentToolExecutor {
  const registry = new AgentToolRegistry()
  const scheduler = new AgentToolScheduler()
  const executor = new AgentToolExecutor({ registry, scheduler })

  executor.registerMany(createWordToolPack(deps.word))
  executor.registerMany(createExcelToolPack(deps.excel))
  executor.registerMany(createPptToolPack(deps.ppt))
  executor.registerMany(createWorkspaceToolPack(deps.workspace))
  executor.registerMany(createWebToolPack(deps.web))
  executor.registerMany(createMemoryToolPack(deps.memory))
  executor.registerMany(createKnowledgeToolPack(deps.knowledge))

  return executor
}
