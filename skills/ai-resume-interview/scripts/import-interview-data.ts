/**
 * Import interview data from JSON files into Supabase
 * Uses service role key to bypass RLS policies
 *
 * Usage: Run from your project root:
 *   npx tsx skills/ai-resume-interview/scripts/import-interview-data.ts
 *
 * Expects:
 *   - .env file in project root with VITE_SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY
 *   - interview-data/ directory in project root with JSON files from the interview
 */

import { createClient } from '@supabase/supabase-js'
import { readFileSync, readdirSync, existsSync } from 'fs'
import { join } from 'path'
import { config } from 'dotenv'

// Use current working directory as project root (run from project root)
const projectRoot = process.cwd()
config({ path: join(projectRoot, '.env') })

// Validate required environment variables
const supabaseUrl = process.env.VITE_SUPABASE_URL
const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY

if (!supabaseUrl || !serviceRoleKey) {
  console.error('Missing required environment variables:')
  if (!supabaseUrl) console.error('  - VITE_SUPABASE_URL')
  if (!serviceRoleKey) console.error('  - SUPABASE_SERVICE_ROLE_KEY')
  console.error('\nAdd these to your .env file')
  process.exit(1)
}

// Create Supabase client with service role key (bypasses RLS)
const supabase = createClient(supabaseUrl, serviceRoleKey, {
  auth: {
    autoRefreshToken: false,
    persistSession: false,
  },
})

const INTERVIEW_DATA_DIR = join(projectRoot, 'interview-data')

interface ImportResult {
  table: string
  action: 'inserted' | 'updated' | 'skipped'
  details: string
}

async function importProfile(data: Record<string, unknown>): Promise<ImportResult> {
  // Check if a profile exists
  const { data: existing } = await supabase
    .from('candidate_profile')
    .select('id')
    .limit(1)
    .single()

  if (existing) {
    // Update existing profile
    const { error } = await supabase
      .from('candidate_profile')
      .update(data)
      .eq('id', existing.id)

    if (error) throw error
    return { table: 'candidate_profile', action: 'updated', details: data.name as string }
  }

  // Insert new profile
  const { error } = await supabase.from('candidate_profile').insert(data)

  if (error) throw error
  return { table: 'candidate_profile', action: 'inserted', details: data.name as string }
}

async function importExperience(data: Record<string, unknown>): Promise<ImportResult> {
  const companyName = data.company_name as string

  // Check if experience exists by company name
  const { data: existing } = await supabase
    .from('experiences')
    .select('id')
    .eq('company_name', companyName)
    .limit(1)
    .single()

  if (existing) {
    // Update existing experience
    const { error } = await supabase
      .from('experiences')
      .update(data)
      .eq('id', existing.id)

    if (error) throw error
    return { table: 'experiences', action: 'updated', details: companyName }
  }

  // Insert new experience
  const { error } = await supabase.from('experiences').insert(data)

  if (error) throw error
  return { table: 'experiences', action: 'inserted', details: companyName }
}

async function importSkills(data: Record<string, unknown>[]): Promise<ImportResult[]> {
  const results: ImportResult[] = []

  for (const skill of data) {
    const skillName = skill.skill_name as string

    // Check if skill exists
    const { data: existing } = await supabase
      .from('skills')
      .select('id')
      .eq('skill_name', skillName)
      .limit(1)
      .single()

    if (existing) {
      const { error } = await supabase.from('skills').update(skill).eq('id', existing.id)
      if (error) throw error
      results.push({ table: 'skills', action: 'updated', details: skillName })
    } else {
      const { error } = await supabase.from('skills').insert(skill)
      if (error) throw error
      results.push({ table: 'skills', action: 'inserted', details: skillName })
    }
  }

  return results
}

async function importGaps(data: Record<string, unknown>[]): Promise<ImportResult[]> {
  const results: ImportResult[] = []

  for (const gap of data) {
    const description = gap.description as string

    // Check if gap exists by description
    const { data: existing } = await supabase
      .from('gaps_weaknesses')
      .select('id')
      .eq('description', description)
      .limit(1)
      .single()

    if (existing) {
      const { error } = await supabase.from('gaps_weaknesses').update(gap).eq('id', existing.id)
      if (error) throw error
      results.push({ table: 'gaps_weaknesses', action: 'updated', details: description.slice(0, 40) })
    } else {
      const { error } = await supabase.from('gaps_weaknesses').insert(gap)
      if (error) throw error
      results.push({ table: 'gaps_weaknesses', action: 'inserted', details: description.slice(0, 40) })
    }
  }

  return results
}

async function importFaqs(data: Record<string, unknown>[]): Promise<ImportResult[]> {
  const results: ImportResult[] = []

  for (const faq of data) {
    const question = faq.question as string

    // Check if FAQ exists by question
    const { data: existing } = await supabase
      .from('faq_responses')
      .select('id')
      .eq('question', question)
      .limit(1)
      .single()

    if (existing) {
      const { error } = await supabase.from('faq_responses').update(faq).eq('id', existing.id)
      if (error) throw error
      results.push({ table: 'faq_responses', action: 'updated', details: question.slice(0, 40) })
    } else {
      const { error } = await supabase.from('faq_responses').insert(faq)
      if (error) throw error
      results.push({ table: 'faq_responses', action: 'inserted', details: question.slice(0, 40) })
    }
  }

  return results
}

async function main() {
  console.log('Importing interview data to Supabase...\n')

  if (!existsSync(INTERVIEW_DATA_DIR)) {
    console.error(`Interview data directory not found: ${INTERVIEW_DATA_DIR}`)
    process.exit(1)
  }

  const files = readdirSync(INTERVIEW_DATA_DIR).filter((f) => f.endsWith('.json'))

  if (files.length === 0) {
    console.log('No JSON files found in interview-data/')
    return
  }

  console.log(`Found ${files.length} JSON files\n`)

  const results: ImportResult[] = []

  for (const file of files) {
    // Skip progress.json - it's for tracking interview state, not data
    if (file === 'progress.json') {
      console.log(`Skipping ${file} (interview progress tracker)`)
      continue
    }

    const filePath = join(INTERVIEW_DATA_DIR, file)
    const content = readFileSync(filePath, 'utf-8')
    const data = JSON.parse(content)

    console.log(`Processing ${file}...`)

    try {
      if (file === 'profile.json') {
        results.push(await importProfile(data))
      } else if (file.startsWith('experience-')) {
        results.push(await importExperience(data))
      } else if (file === 'skills.json') {
        results.push(...(await importSkills(Array.isArray(data) ? data : [data])))
      } else if (file === 'gaps.json') {
        results.push(...(await importGaps(Array.isArray(data) ? data : [data])))
      } else if (file === 'faqs.json') {
        results.push(...(await importFaqs(Array.isArray(data) ? data : [data])))
      } else {
        console.log(`  Unknown file type, skipping`)
      }
    } catch (error) {
      console.error(`  Error processing ${file}:`, error)
    }
  }

  // Summary
  console.log('\n--- Import Summary ---')
  const inserted = results.filter((r) => r.action === 'inserted')
  const updated = results.filter((r) => r.action === 'updated')

  if (inserted.length > 0) {
    console.log(`\nInserted (${inserted.length}):`)
    inserted.forEach((r) => console.log(`  ${r.table}: ${r.details}`))
  }

  if (updated.length > 0) {
    console.log(`\nUpdated (${updated.length}):`)
    updated.forEach((r) => console.log(`  ${r.table}: ${r.details}`))
  }

  console.log(`\nTotal: ${results.length} records processed`)
}

main().catch((error) => {
  console.error('Import failed:', error)
  process.exit(1)
})
