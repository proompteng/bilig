function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

const currentStarterIssueNumbers = [273, 283, 285, 300, 334, 358, 360, 361, 362, 363, 366, 367, 368, 369, 371] as const

const closedStarterIssueNumbers = [
  137, 138, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 154, 224, 231, 199, 200, 201, 202, 203, 204, 205, 228, 229, 246,
  266, 282, 294, 160, 161, 164, 165, 166, 168, 169, 170, 171, 172, 173, 174, 175, 176, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187,
  188, 189, 190, 191, 192, 276, 227, 247, 256, 315, 316, 317, 318, 319, 336, 341, 343, 344, 345, 346, 347, 354, 364, 374, 377, 380,
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
