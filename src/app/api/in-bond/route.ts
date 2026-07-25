import { NextRequest, NextResponse } from 'next/server';
import { cookies } from 'next/headers';
import supabase from '@/lib/db';

export const dynamic = 'force-dynamic';

function getUser() {
  const cookieStore = cookies();
  return cookieStore.get('skyroute_user')?.value || null;
}

export async function POST(req: NextRequest) {
  const user = getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  try {
    const body = await req.json();
    const direction = body.direction === 'outbound' ? 'outbound' : 'inbound';

    const { data, error } = await supabase
      .from('in_bond_sheets')
      .insert([{ ...body, direction, created_by: user }])
      .select()
      .single();

    if (error) throw error;
    return NextResponse.json({ success: true, data });
  } catch (err: any) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

export async function GET(req: NextRequest) {
  const user = getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  try {
    const { searchParams } = new URL(req.url);
    const direction = searchParams.get('direction');
    const search = searchParams.get('search');
    const limit = Number(searchParams.get('limit')) || 100;

    let query = supabase
      .from('in_bond_sheets')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(limit);

    if (direction === 'inbound' || direction === 'outbound') {
      query = query.eq('direction', direction);
    }

    if (search) {
      query = query.or(
        `bar_number.ilike.%${search}%,c209_number.ilike.%${search}%,flight_number.ilike.%${search}%,container_code.ilike.%${search}%`
      );
    }

    const { data, error } = await query;
    if (error) throw error;
    return NextResponse.json({ data });
  } catch (err: any) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

export async function PATCH(req: NextRequest) {
  const user = getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  try {
    const body = await req.json();
    const { id, ...updates } = body;
    if (!id) return NextResponse.json({ error: 'Missing id' }, { status: 400 });

    const { data, error } = await supabase
      .from('in_bond_sheets')
      .update({ ...updates, updated_by: user, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;
    return NextResponse.json({ success: true, data });
  } catch (err: any) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

export async function DELETE(req: NextRequest) {
  const user = getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  try {
    const { searchParams } = new URL(req.url);
    const id = searchParams.get('id');
    if (!id) return NextResponse.json({ error: 'Missing id' }, { status: 400 });

    const { error } = await supabase.from('in_bond_sheets').delete().eq('id', id);
    if (error) throw error;
    return NextResponse.json({ success: true });
  } catch (err: any) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}
