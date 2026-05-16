import type { CommunityGrowthSnapshot, GitHubDiscussionActivitySnapshot, GitHubTrafficSnapshot } from './community-growth-snapshot.ts'

const starGoal = 1000

function formatCount(value: number): string {
  return String(value).replace(/\B(?=(\d{3})+(?!\d))/g, ',')
}

function markdownLink(label: string, url: string): string {
  return `[${label}](${url})`
}

function formatCommentCount(value: number): string {
  return `${formatCount(value)} ${value === 1 ? 'comment' : 'comments'}`
}

function renderDiscussionActivityMarkdown(discussionActivity: GitHubDiscussionActivitySnapshot): readonly string[] {
  if (!discussionActivity.available) {
    return [`- Discussion activity: unavailable. ${discussionActivity.reason}`]
  }

  const lines = [`- Total discussions: ${formatCount(discussionActivity.totalCount)}`]
  for (const discussion of discussionActivity.recent.slice(0, 5)) {
    lines.push(
      `- #${String(discussion.number)} ${markdownLink(discussion.title, discussion.url)} (${discussion.category}, ${formatCommentCount(discussion.commentCount)})`,
    )
  }
  return lines
}

function renderTrafficMarkdown(traffic: GitHubTrafficSnapshot): readonly string[] {
  if (!traffic.available) {
    return [`- Traffic: unavailable. ${traffic.reason}`]
  }

  const lines = [
    `- Views: ${formatCount(traffic.views.count)} from ${formatCount(traffic.views.uniques)} unique visitors`,
    `- Clones: ${formatCount(traffic.clones.count)} from ${formatCount(traffic.clones.uniques)} unique cloners`,
  ]

  if (traffic.referrers.length > 0) {
    lines.push(
      `- Top referrers: ${traffic.referrers
        .slice(0, 5)
        .map((referrer) => `${referrer.referrer} (${formatCount(referrer.count)}/${formatCount(referrer.uniques)})`)
        .join(', ')}`,
    )
  }

  if (traffic.paths.length > 0) {
    lines.push(
      `- Top paths: ${traffic.paths
        .slice(0, 5)
        .map((path) => `${path.path} (${formatCount(path.count)}/${formatCount(path.uniques)})`)
        .join(', ')}`,
    )
  }

  return lines
}

function ratioPerStar(value: number, stars: number): string {
  if (stars <= 0) {
    return 'n/a'
  }
  return formatCount(Math.round(value / stars))
}

function renderConversionPressureMarkdown(snapshot: CommunityGrowthSnapshot): readonly string[] {
  const stars = snapshot.github.stargazerCount
  const lines = [
    `- Last-week npm downloads per current star: ${ratioPerStar(snapshot.npm.downloads.lastWeek.downloads, stars)}`,
    `- Last-month npm downloads per current star: ${ratioPerStar(snapshot.npm.downloads.lastMonth.downloads, stars)}`,
  ]

  if (snapshot.traffic.available) {
    lines.push(
      `- Fourteen-day unique GitHub visitors per current star: ${ratioPerStar(snapshot.traffic.views.uniques, stars)}`,
      `- Fourteen-day unique cloners per current star: ${ratioPerStar(snapshot.traffic.clones.uniques, stars)}`,
    )
  } else {
    lines.push(`- GitHub traffic pressure: unavailable. ${snapshot.traffic.reason}`)
  }

  lines.push(
    '- Interpretation: these are pressure ratios, not attribution. High download or clone pressure with flat stars means the evaluator path needs a clearer proof, trust signal, or bookmark ask after verification.',
  )

  return lines
}

function renderSpikeReadMarkdown(snapshot: CommunityGrowthSnapshot): readonly string[] {
  if (!snapshot.traffic.available) {
    return [
      '- Traffic referrers are unavailable in this snapshot, so the spike read cannot be refreshed from live GitHub traffic.',
      '- Keep the next distribution move narrow: one proof URL, one adoption-blocker question, and a before/after traffic snapshot.',
    ]
  }

  const topExternal = snapshot.traffic.referrers.find((referrer) => referrer.referrer !== 'github.com')
  const hackerNews = snapshot.traffic.referrers.find((referrer) => referrer.referrer === 'news.ycombinator.com')
  const twitter = snapshot.traffic.referrers.find((referrer) => referrer.referrer === 't.co')

  const lines = ['- The visible May 7-11 star jump still lines up with external developer traffic, not broad social posting.']

  if (topExternal !== undefined) {
    lines.push(
      `- The strongest current external referrer is ${topExternal.referrer} with ${formatCount(topExternal.count)} views from ${formatCount(
        topExternal.uniques,
      )} unique visitors.`,
    )
  }

  if (hackerNews !== undefined && twitter !== undefined) {
    lines.push(
      `- Hacker News is still ahead of X/t.co in qualified GitHub traffic: ${formatCount(hackerNews.count)}/${formatCount(
        hackerNews.uniques,
      )} versus ${formatCount(twitter.count)}/${formatCount(twitter.uniques)}.`,
    )
  }

  lines.push(
    '- Replication plan: do not repost the same launch. Ship one sharper proof page, then ask HN/X/MCP audiences for a concrete adoption blocker: formula family, XLSX cache behavior, persistence shape, or agent writeback verification.',
  )

  return lines
}

export function renderCommunityGrowthSnapshotMarkdown(snapshot: CommunityGrowthSnapshot): string {
  const starRemaining = Math.max(0, starGoal - snapshot.github.stargazerCount)
  const lines = [
    '# Community Growth Snapshot',
    '',
    `Captured at: \`${snapshot.capturedAt}\``,
    '',
    'This snapshot tracks the public signals for the `@bilig/headless` growth loop: GitHub conversion, npm demand, contributor on-ramp health, discussion activity, and traffic quality.',
    '',
    '## GitHub',
    '',
    `- Repository: ${markdownLink(snapshot.github.fullName, snapshot.github.htmlUrl)}`,
    `- Stars: ${formatCount(snapshot.github.stargazerCount)} / ${formatCount(starGoal)} (${formatCount(starRemaining)} remaining)`,
    `- Forks: ${formatCount(snapshot.github.forkCount)}`,
    `- Watchers: ${formatCount(snapshot.github.watcherCount)}`,
    `- Open issues: ${formatCount(snapshot.github.openIssueCount)}`,
    `- Default branch: \`${snapshot.github.defaultBranch}\``,
    `- Topics: ${snapshot.github.topics.map((topic) => `\`${topic}\``).join(', ') || 'none'}`,
    '',
    '## npm',
    '',
    `- Package: \`${snapshot.npm.name}@${snapshot.npm.version}\``,
    `- License: \`${snapshot.npm.license || 'unknown'}\``,
    `- Modified: \`${snapshot.npm.modifiedAt}\``,
    `- Downloads last week: ${formatCount(snapshot.npm.downloads.lastWeek.downloads)} (${snapshot.npm.downloads.lastWeek.start} to ${snapshot.npm.downloads.lastWeek.end})`,
    `- Downloads last month: ${formatCount(snapshot.npm.downloads.lastMonth.downloads)} (${snapshot.npm.downloads.lastMonth.start} to ${snapshot.npm.downloads.lastMonth.end})`,
    '',
    '## Contributor Funnel',
    '',
    `- Open good first issues: ${formatCount(snapshot.contributorFunnel.openGoodFirstIssueCount)}`,
    `- Open first-timers-only issues: ${formatCount(snapshot.contributorFunnel.openFirstTimersOnlyIssueCount)}`,
    `- Documentation starter issues: ${formatCount(snapshot.contributorFunnel.openDocumentationStarterIssueCount)}`,
    `- Non-documentation starter issues: ${formatCount(snapshot.contributorFunnel.openNonDocumentationStarterIssueCount)}`,
    `- Open help wanted issues: ${formatCount(snapshot.contributorFunnel.openHelpWantedIssueCount)}`,
    `- Open pull requests: ${formatCount(snapshot.contributorFunnel.openPullRequestCount)}`,
    `- External open issues: ${formatCount(snapshot.contributorFunnel.externalOpenIssueCount)}`,
    `- External open pull requests: ${formatCount(snapshot.contributorFunnel.externalOpenPullRequestCount)}`,
    `- External issues opened in the last 7 days: ${formatCount(snapshot.contributorFunnel.externalIssuesOpenedLastSevenDays)}`,
    `- External pull requests opened in the last 7 days: ${formatCount(snapshot.contributorFunnel.externalPullRequestsOpenedLastSevenDays)}`,
    '',
    '## Discussions',
    '',
    ...renderDiscussionActivityMarkdown(snapshot.discussionActivity),
    '',
    '## Traffic',
    '',
    ...renderTrafficMarkdown(snapshot.traffic),
    '',
    '## Conversion Pressure',
    '',
    ...renderConversionPressureMarkdown(snapshot),
    '',
    '## Spike Read',
    '',
    ...renderSpikeReadMarkdown(snapshot),
    '',
    '## Read This Snapshot',
    '',
    '- Stars are the primary lagging goal; npm downloads, clone traffic, and external issues are leading signals.',
    '- If downloads or clones rise without stars, improve README and npm star/bookmark conversion after proof blocks.',
    '- If traffic comes from a channel but discussions stay quiet, switch from launch copy to a specific workflow-feedback ask.',
    '- If the starter queue drops below three current issues, open scoped example tasks before the next distribution push.',
  ]

  return `${lines.join('\n')}\n`
}
