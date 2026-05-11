import { NextResponse } from "next/server"
import Anthropic from "@anthropic-ai/sdk"

const apiKey = process.env.CLAUDE_API_KEY || process.env.ANTHROPIC_API_KEY
const client = new Anthropic({ apiKey })

export async function POST(request) {
  try {
    const { system, messages, max_tokens } = await request.json()

    if (!apiKey) {
      return NextResponse.json(
        { error: "API key is not configured" },
        { status: 500 }
      )
    }

    const response = await client.messages.create({
      model: "claude-sonnet-4-20250514",
      max_tokens: max_tokens || 1800,
      system,
      messages,
    })

    return NextResponse.json(response)
  } catch (err) {
    console.error("Claude API error:", err)
    return NextResponse.json(
      { error: err.message || "Claude API request failed" },
      { status: 500 }
    )
  }
}
