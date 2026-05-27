import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://weqrtoovivzbunaakpca.supabase.co/rest/v1/';
const supabaseKey = 'sb_publishable_5wCORNNAyljwtOZHY67oew_4iQz1o4_';

export const supabase = createClient(
  supabaseUrl,
  supabaseKey
);
