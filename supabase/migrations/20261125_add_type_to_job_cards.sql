-- Migration: Add type column to job_cards table
-- Options: Drawing, Evatek, Tender, Modeling

ALTER TABLE job_cards 
  ADD COLUMN IF NOT EXISTS type text DEFAULT 'Drawing';
