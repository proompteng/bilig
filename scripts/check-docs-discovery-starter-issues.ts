function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

const currentStarterIssueNumbers = [
  134, 153, 155, 156, 158, 159, 162, 163, 193, 194, 195, 196, 197, 198, 207, 208, 209, 210, 211, 212, 217, 218, 219, 220, 221, 222, 223,
  233, 248, 249, 250, 255, 256, 257, 258, 259, 260, 265, 267, 268, 269, 272, 273, 274, 275, 277, 278, 279, 280, 281, 283, 284, 285, 286,
  287, 288, 289, 290, 292, 293, 296, 297, 298, 299, 300, 301, 302, 303, 304, 305, 306, 309, 310, 311, 312, 313, 314, 323, 324, 325, 326,
  327, 328, 329, 330, 331, 332, 333, 334,
] as const

const closedStarterIssueNumbers = [
  137, 138, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 154, 224, 231, 199, 200, 201, 202, 203, 204, 205, 228, 229, 246,
  266, 282, 294, 160, 161, 164, 165, 166, 168, 169, 170, 171, 172, 173, 174, 175, 176, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187,
  188, 189, 190, 191, 192, 276, 227, 247, 315, 316, 317, 318, 319, 336,
] as const

export function requireStarterIssueDiscovery(starterIssues: string, llms: string): void {
  for (const issueNumber of currentStarterIssueNumbers) {
    const required = `https://github.com/proompteng/bilig/issues/${issueNumber}`
    requireIncludes(starterIssues, required, 'docs/starter-issues.md')
    requireIncludes(llms, required, 'docs/llms.txt')
  }

  for (const issueNumber of closedStarterIssueNumbers) {
    const issueUrl = `https://github.com/proompteng/bilig/issues/${issueNumber}`
    if (starterIssues.includes(issueUrl)) {
      throw new Error(`docs/starter-issues.md still links to closed starter issue #${issueNumber}`)
    }

    if (llms.includes(issueUrl)) {
      throw new Error(`docs/llms.txt still links to closed starter issue #${issueNumber}`)
    }
  }
}
