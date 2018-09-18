CREATE TABLE IF NOT EXISTS "Track"(
  "title" VARCHAR(128) NOT NULL,
  "album" VARCHAR(128) NOT NULL,
  "artist" VARCHAR(128) NOT NULL,
  "genre" VARCHAR(64),
  "lenght" INTEGER,
  "rating" INTEGER,
  PRIMARY KEY("title","album","artist")
);