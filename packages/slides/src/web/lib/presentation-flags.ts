/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Front-end feature flags for the presentation editor.
 *
 * General-release interaction flags are default-ON and accept the exact value
 * `"0"` as a deploy-time kill switch.
 */








export function presentationMultiTransformEnabled(): boolean {
  return process.env.NEXT_PUBLIC_PRESENTATION_MULTI_TRANSFORM_V1 !== "0";
}






export function presentationPressDragEnabled(): boolean {
  return process.env.NEXT_PUBLIC_PRESENTATION_PRESS_DRAG_V1 !== "0";
}








export function presentationRailVirtualEnabled(): boolean {
  return process.env.NEXT_PUBLIC_PRESENTATION_RAIL_VIRTUAL_V1 !== "0";
}






export function presentationViewportPanEnabled(): boolean {
  return process.env.NEXT_PUBLIC_PRESENTATION_VIEWPORT_PAN_V1 !== "0";
}







export function presentationHistoryCoalesceEnabled(): boolean {
  return process.env.NEXT_PUBLIC_PRESENTATION_HISTORY_COALESCE_V1 !== "0";
}










export function presentationConnectorAnchorsEnabled(): boolean {
  return process.env.NEXT_PUBLIC_PRESENTATION_CONNECTOR_ANCHORS_V1 !== "0";
}










export function presentationElbowConnectorEnabled(): boolean {
  return process.env.NEXT_PUBLIC_PRESENTATION_ELBOW_CONNECTOR_V1 !== "0";
}
