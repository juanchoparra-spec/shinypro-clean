import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = 'https://xhmcroyrjkgbcpwahyia.supabase.co'
const SUPABASE_KEY = 'sb_publishable_5TAyOpJwK4K7gkB4nfhPzQ_TtaHXBw-'

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY)
